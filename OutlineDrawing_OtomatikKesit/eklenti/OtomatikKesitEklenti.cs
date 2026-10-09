// =====================================================================
//  OtomatikKesit.dll  -  OutlineDrawing için ayrı eklenti
// ---------------------------------------------------------------------
//  OutlineDrawing.dll'e dokunmaz. İkisi de NETLOAD ile yüklenir; OutlineDrawing
//  formu açıldığında "Kesit Alma" kutusuna şunlar eklenir:
//    - "Oto Kesit" butonu: .tow dosyasını OutlineDrawing'in kendi okuyucusuyla
//      (dataprocessing.Run) okur; ön görünüş (frontFace) ve yan görünüş (sideFace)
//      dışında kalan elemanların hepsi bir kesite girecek şekilde kesitleri bulur
//      ve dataGridView1'e yazar. Kapsanamayan eleman kalırsa komut satırında listeler.
//    - "Ölçü" kutucuğu: işaretliyse ÇALIŞTIR'dan sonra her "SECTION x" çizimine
//      ölçü koyar (üstte zincir ölçü, sağda toplam derinlik).
//    - "Redundant" kutucuğu: işaretliyse ÇALIŞTIR'dan sonra ön/yan görünüşte ve kesitlerde
//      ana elemanlar düz/cyan, redundant elemanlar kesikli/mavi olur. Redundant: PLS-TOWER grup tipi
//      "Redundant" (grup tablosunda tip 3, PLS'te turuncu) veya grup açıklamasında "Redundant" geçen gruplar.
//
//  OutlineDrawing'in iç tiplerine yansıtma (reflection) ile erişilir, bu yüzden
//  OutlineDrawing.dll'in yeniden derlenmesi gerekmez.
//  Komutlar: OTOKESIT (tabloyu doldur), OTOOLCU (çizili kesitlere ölçü koy / yenile),
//            OTOSTIL (kesitlerde redundant elemanları kesikli yap)
// =====================================================================

using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
using AcColor = Autodesk.AutoCAD.Colors.Color;

[assembly: ExtensionApplication(typeof(OtomatikKesit.Eklenti))]
#if !TEK_DLL
// Tek DLL (OutlineDrawing ile birleştirilmiş) derlemede bu satır olmamalı: assembly'de CommandClass varsa
// AutoCAD yalnız listelenen sınıfları tarar ve OutlineDrawing'in DRAWOUTLINE / CONVERTO2D komutları kaybolur.
[assembly: CommandClass(typeof(OtomatikKesit.Komutlar))]
#endif

namespace OtomatikKesit
{
    public class Eklenti : IExtensionApplication
    {
        public const string Surum = "1.13";
        private static Timer _timer;

        public void Initialize()
        {
            // Modeless OutlineDrawing formunu yakalamak için UI iş parçacığında küçük bir zamanlayıcı
            _timer = new Timer { Interval = 700 };
            _timer.Tick += (s, e) =>
            {
                try { FormBaglayici.FormlariTara(); }
                catch { }
            };
            _timer.Start();
            Komutlar.Yaz("\nOtomatik Kesit " + Surum + " yüklendi. OutlineDrawing formunda 'Oto Kesit' butonu görünecek (komutlar: OTOKESIT, OTOOLCU, OTOSTIL).\n");
        }

        public void Terminate()
        {
            if (_timer != null) _timer.Stop();
        }
    }

    public class Komutlar
    {
        [CommandMethod("OTOKESIT", CommandFlags.Session)]
        public void OtoKesit()
        {
            Form f = FormBaglayici.AcikFormuBul();
            if (f == null)
            {
                Yaz("\nOutlineDrawing formu açık değil. Önce OutlineDrawing'i çalıştırıp .tow dosyasını seçin.\n");
                return;
            }
            FormBaglayici.TabloyuDoldur(f);
        }

        [CommandMethod("OTOOLCU", CommandFlags.Session)]
        public void OtoOlcu()
        {
            FormBaglayici.Guvenli(null, "Ölçü", () => KesitOlculeri.Ekle());
        }

        [CommandMethod("OTOSTIL", CommandFlags.Session)]
        public void OtoStil()
        {
            Form f = FormBaglayici.AcikFormuBul();
            if (f == null)
            {
                Yaz("\nOutlineDrawing formu açık değil (kesit tanımları ve .tow yolu formdan okunur).\n");
                return;
            }
            FormBaglayici.Guvenli(f, "Çizgi tipi", () => Yaz("\n" + KesitStili.Uygula(f) + "\n"));
        }

        internal static void Yaz(string msg)
        {
            try
            {
                var doc = AcApp.DocumentManager.MdiActiveDocument;
                if (doc != null) doc.Editor.WriteMessage(msg);
            }
            catch { }
        }
    }

    internal static class FormBaglayici
    {
        private const string FormTipi = "OutlineDrawing.mainform";
        private const string ButonAdi = "btnAutoSection";
        private const string OlcuKutusuAdi = "chkAutoDim";
        private const BindingFlags Hepsi = BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;
        private static readonly string[] NoktaSutunlari = { "ColumnFirstPoint", "ColumnSecondPoint", "ColumnThirdPoint", "ColumnFourthPoint" };

        internal static Form AcikFormuBul()
        {
            foreach (Form f in System.Windows.Forms.Application.OpenForms)
                if (!f.IsDisposed && f.GetType().FullName == FormTipi) return f;
            return null;
        }

        internal static void FormlariTara()
        {
            Form f = AcikFormuBul();
            if (f == null || f.Controls.Find(ButonAdi, true).Length > 0) return;
            KontrolleriEkle(f);
        }

        private static void KontrolleriEkle(Form f)
        {
            var ornek = f.Controls.Find("btnHeight", true).FirstOrDefault() as Button;
            Control kap = ornek != null ? ornek.Parent : (Control)f;

            var b = new Button { Name = ButonAdi, Text = "Oto Kesit" };
            var chk = new CheckBox { Name = OlcuKutusuAdi, Text = "Ölçü", Checked = true, AutoSize = true };
            var chkR = new CheckBox { Name = "chkRedundant", Text = "Redundant", Checked = true, AutoSize = true };
            if (ornek != null)
            {
                b.Size = ornek.Size;
                b.Font = ornek.Font;
                b.FlatStyle = ornek.FlatStyle;
                b.BackColor = ornek.BackColor;
                b.ForeColor = ornek.ForeColor;
                b.UseVisualStyleBackColor = ornek.UseVisualStyleBackColor;
                b.Cursor = ornek.Cursor;
                b.Anchor = ornek.Anchor;
                b.FlatAppearance.BorderColor = ornek.FlatAppearance.BorderColor;
                b.FlatAppearance.BorderSize = ornek.FlatAppearance.BorderSize;
                // Aynı sütundaki (Nokta, Yükseklik, Temizle, 2D Aktar, Liste) son butonun altına
                int alt = kap.Controls.Cast<Control>()
                             .Where(c => c is Button && Math.Abs(c.Left - ornek.Left) <= 3)
                             .Max(c => c.Bottom);
                b.Location = new Point(ornek.Left, alt + 6);
                chk.Font = ornek.Font;
                chk.ForeColor = kap.ForeColor;
                chk.Anchor = ornek.Anchor;
                chk.Location = new Point(ornek.Left + 2, b.Bottom + 4);
                chkR.Font = ornek.Font;
                chkR.ForeColor = kap.ForeColor;
                chkR.Anchor = ornek.Anchor;
                chkR.Location = new Point(ornek.Left + 2, chk.Bottom + 2);
            }
            else
            {
                b.Size = new Size(70, 24);
                b.Location = new Point(8, 8);
                chk.Location = new Point(8, 36);
                chkR.Location = new Point(8, 58);
            }

            kap.Controls.Add(b);
            kap.Controls.Add(chk);
            kap.Controls.Add(chkR);
            b.BringToFront();
            chk.BringToFront();
            chkR.BringToFront();
            var tip = new ToolTip();
            tip.SetToolTip(b, "Ön/yan görünüşte görünmeyen elemanların hepsini kapsayan kesitleri bulup tabloya yazar");
            tip.SetToolTip(chk, "ÇALIŞTIR'dan sonra kesitlere ölçü koy (komut: OTOOLCU)");
            tip.SetToolTip(chkR, "YÜKLE ve ÇALIŞTIR'dan sonra 3D modelde, görünüşlerde ve kesitlerde ana elemanlar düz, redundant elemanlar kesikli; isimler tek renk (komut: OTOSTIL)");
            b.Click += (s, e) => Guvenli(f, "Otomatik kesit", () => TabloyuDoldur(f));

            // YÜKLE ve ÇALIŞTIR model alanındaki HER ŞEYİ siler: başka katmanda çizim varsa önce sor
            var load = Alan(f, "btnLoad") as Button;
            var run = Alan(f, "btnRun") as Button;
            SilmeKorumasiEkle(f, load, "YÜKLE");
            SilmeKorumasiEkle(f, run, "ÇALIŞTIR");

            // YÜKLE (3D model) sonrasında: redundant kesikli + .tow kontrolü
            if (load != null)
                load.Click += (s, e) => Guvenli(f, "YÜKLE sonrası", () =>
                {
                    var mesaj = new List<string>();
                    string uyari = TowKontrol.Uyari(YolAl(f));
                    if (uyari != null) mesaj.Add(uyari);
                    if (chkR.Checked) mesaj.Add(KesitStili.Uygula(f));
                    if (mesaj.Count > 0) Uyari(f, string.Join(" | ", mesaj));
                });

            // ÇALIŞTIR'a ikinci bir olay bağlanır; OutlineDrawing'in kendi çizimi bittikten sonra çalışır.
            if (run != null)
                run.Click += (s, e) => Guvenli(f, "ÇALIŞTIR sonrası", () =>
                {
                    var mesaj = new List<string>();
                    if (chkR.Checked)
                    {
                        try { mesaj.Add(KesitStili.Uygula(f)); }
                        catch (System.Exception ex) { mesaj.Add("Redundant çizgi tipi uygulanamadı: " + IcHata(ex).Message); }
                    }
                    if (chk.Checked)
                    {
                        try { mesaj.Add(KesitOlculeri.Ekle()); }
                        catch (System.Exception ex) { mesaj.Add("Ölçü eklenemedi: " + IcHata(ex).Message); }
                    }
                    if (mesaj.Count > 0) Uyari(f, string.Join(" | ", mesaj));
                });
        }

        // ---------------- güvenlik yardımcıları ----------------

        internal static System.Exception IcHata(System.Exception ex)
        {
            while (ex is TargetInvocationException && ex.InnerException != null) ex = ex.InnerException;
            return ex;
        }

        /// <summary>Olay/komut gövdesini çalıştırır; hiçbir hata AutoCAD'e sızmaz.</summary>
        internal static void Guvenli(Form f, string ne, Action a)
        {
            try { a(); }
            catch (System.Exception ex)
            {
                System.Exception ic = IcHata(ex);
                Komutlar.Yaz("\n" + ne + " hatası: " + ic.Message + "\n");
                try
                {
                    if (f != null && !f.IsDisposed)
                        MessageBox.Show(f, ne + " sırasında hata:\n" + ic.Message, "Otomatik Kesit", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch { }
            }
        }

        internal static string YolAl(Form f)
        {
            var txt = Alan(f, "txtTowerPath") as TextBox;
            return txt != null ? (txt.Text ?? "").Trim() : "";
        }

        private static bool _silmeOnaylandi;

        /// <summary>Butonun mevcut (OutlineDrawing) Click olayını, önce silme kontrolü yapan bir olayın içine alır.
        /// OutlineDrawing'in runtowerfile / loadtowerfile'ı model alanındaki tüm nesneleri siliyor.</summary>
        private static void SilmeKorumasiEkle(Form f, Button btn, string ad)
        {
            if (btn == null) return;
            try
            {
                PropertyInfo ep = typeof(System.ComponentModel.Component).GetProperty("Events", BindingFlags.Instance | BindingFlags.NonPublic);
                FieldInfo kf = typeof(Control).GetField("EventClick", BindingFlags.Static | BindingFlags.NonPublic);
                var liste = ep != null ? ep.GetValue(btn, null) as System.ComponentModel.EventHandlerList : null;
                object anahtar = kf != null ? kf.GetValue(null) : null;
                if (liste == null || anahtar == null) return;
                var asil = liste[anahtar] as EventHandler;
                if (asil == null) return;
                liste.RemoveHandler(anahtar, asil);
                btn.Click += (s, e) =>
                {
                    bool devam = true;
                    try { devam = SilmeOnayi(f, ad); } catch { devam = true; }
                    if (devam) asil(s, e);
                };
            }
            catch { /* olaylar okunamazsa koruma eklenmez, OutlineDrawing olduğu gibi çalışır */ }
        }

        private static bool SilmeOnayi(Form f, string ad)
        {
            if (_silmeOnaylandi) return true;
            int n = YabanciNesneler.Say();
            if (n == 0) return true;
            DialogResult c = MessageBox.Show(f,
                string.Format("Model alanında OutlineDrawing'e ait olmayan {0} nesne var (başka katmanlarda).\n\n" +
                              "{1}, model alanındaki HER ŞEYİ silip yeniden çizer. Devam edilsin mi?\n\n" +
                              "(Evet derseniz bu oturumda bir daha sorulmaz.)", n, ad),
                "Otomatik Kesit - silme uyarısı", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (c != DialogResult.Yes) return false;
            _silmeOnaylandi = true;
            return true;
        }

        internal static void TabloyuDoldur(Form f)
        {
            try
            {
                var txt = Alan(f, "txtTowerPath") as TextBox;
                string path = txt != null ? (txt.Text ?? "").Trim() : "";
                if (path.Length == 0 || !File.Exists(path))
                {
                    Uyari(f, "Önce tower dosyası (.tow) seçin.");
                    return;
                }

                KuleVerisi kule = KuleOkuyucu.Oku(f.GetType().Assembly, path);
                if (kule == null || kule.Uyeler.Count == 0)
                {
                    Uyari(f, "Tower dosyası okunamadı.");
                    return;
                }

                List<AutoSectionCandidate> bulunan;
                AutoSectionReport rapor = null;
                if (kule.GorunenAnahtarlar != null && kule.GorunenAnahtarlar.Count > 0)
                    bulunan = AutoSectionDetector.DetectHidden(kule.Uyeler, kule.GorunenAnahtarlar, new AutoSectionOptions(), out rapor);
                else
                    bulunan = AutoSectionDetector.Detect(kule.Uyeler, new AutoSectionOptions());

                if (bulunan.Count == 0)
                {
                    Uyari(f, rapor != null && rapor.HiddenMembers == 0
                        ? "Bütün elemanlar ön/yan görünüşte görünüyor; kesit gerekmiyor."
                        : "Otomatik kesit bulunamadı.");
                    return;
                }

                var grid = Alan(f, "dataGridView1") as DataGridView;
                if (grid == null)
                {
                    MessageBox.Show(f, "Kesit tablosu (dataGridView1) bulunamadı.", "Otomatik Kesit");
                    return;
                }

                bool doluSatirVar = grid.Rows.Cast<DataGridViewRow>().Any(r => !r.IsNewRow &&
                    r.Cells.Cast<DataGridViewCell>().Any(c => c.OwningColumn.Name != "ColumnSection" &&
                                                              c.Value != null && c.Value.ToString().Trim().Length > 0));
                if (doluSatirVar)
                {
                    DialogResult cevap = MessageBox.Show(f,
                        "Tablodaki mevcut kesitler silinsin mi?\n\nEvet = sil, otomatik kesitleri yaz\nHayır = mevcutların altına ekle",
                        "Otomatik Kesit", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    if (cevap == DialogResult.Cancel) return;
                    if (cevap == DialogResult.Yes) grid.Rows.Clear();
                }

                var satirlar = new List<string>();
                foreach (AutoSectionCandidate c in bulunan)
                {
                    int idx = grid.Rows.Add();
                    DataGridViewRow row = grid.Rows[idx];
                    if (c.Kind == AutoSectionKind.Height)
                    {
                        Hucre(row, "ColumnHeight", c.Height.Value.ToString("0.###"));
                    }
                    else
                    {
                        for (int k = 0; k < NoktaSutunlari.Length && k < c.Points.Count; k++)
                            Hucre(row, NoktaSutunlari[k], AutoSectionDetector.FormatPoint(c.Points[k]));
                    }
                    if (grid.Columns.Contains("ColumnSection"))
                        row.Cells["ColumnSection"].ToolTipText = c.Description;
                    satirlar.Add(c.Description);
                }

                Cagir(f, "RefreshSectionLetters");

                string ozet = string.Format("{0} kesit eklendi", bulunan.Count);
                if (rapor != null)
                    ozet += string.Format(" (görünmeyen {0}/{1} eleman kapsandı{2})", rapor.CoveredHidden, rapor.HiddenMembers,
                        rapor.UncoveredKeys.Count > 0 ? ", liste komut satırında" : "");
                string towUyari = TowKontrol.Uyari(path);
                Uyari(f, ozet + ". Kontrol edip ÇALIŞTIR'a basın." + (towUyari != null ? " " + towUyari : ""));

                string log = "\nOtomatik kesitler:\n  " + string.Join("\n  ", satirlar) + "\n";
                if (rapor != null)
                {
                    log += string.Format("Ön/yan görünüşte görünmeyen eleman: {0}, kesitlerle kapsanan: {1}, kapsanmayan: {2}\n",
                        rapor.HiddenMembers, rapor.CoveredHidden, rapor.UncoveredKeys.Count);
                    foreach (string d in rapor.UncoveredDescriptions) log += "  KAPSANMAYAN: " + d + "\n";
                }
                Komutlar.Yaz(log);
            }
            catch (System.Exception ex)
            {
                System.Exception ic = ex is TargetInvocationException && ex.InnerException != null ? ex.InnerException : ex;
                MessageBox.Show(f, "Otomatik kesit sırasında hata:\n" + ic.Message, "Otomatik Kesit",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void Hucre(DataGridViewRow row, string sutun, string deger)
        {
            if (row.DataGridView.Columns.Contains(sutun)) row.Cells[sutun].Value = deger;
        }

        private static ToolTip _uyariIpucu;

        /// <summary>OutlineDrawing'in uyarı etiketi (lblError) tek satırda sağdaki seçenek kutusunun altında kalıyordu:
        /// genişlik sınırı verilir (metin alt satıra kayar) ve tam metin ipucu balonunda gösterilir.</summary>
        private static void UyariEtiketiniHazirla(Form f, string mesaj)
        {
            var lbl = Alan(f, "lblError") as Label;
            if (lbl == null || lbl.Parent == null) return;
            int sag = lbl.Parent.ClientSize.Width - 8;
            foreach (Control c in lbl.Parent.Controls)
                if (c != lbl && c.Visible && c.Left > lbl.Left && c.Top < lbl.Bottom + 60 && c.Bottom > lbl.Top)
                    sag = Math.Min(sag, c.Left - 8);
            int gen = Math.Max(150, sag - lbl.Left);
            if (lbl.MaximumSize.Width != gen) lbl.MaximumSize = new Size(gen, 0);
            if (_uyariIpucu == null) _uyariIpucu = new ToolTip { AutoPopDelay = 30000 };
            _uyariIpucu.SetToolTip(lbl, mesaj);
        }

        internal static void Uyari(Form f, string mesaj)
        {
            try { UyariEtiketiniHazirla(f, mesaj); } catch { }
            // OutlineDrawing'in kendi uyarı etiketi; yoksa komut satırı
            bool oldu = false;
            try { oldu = Cagir(f, "PrintErrorLabel", 8000, mesaj); } catch { }
            if (!oldu) Komutlar.Yaz("\n" + mesaj + "\n");
        }

        internal static object Alan(object o, string ad)
        {
            FieldInfo fi = o.GetType().GetField(ad, Hepsi);
            return fi != null ? fi.GetValue(o) : null;
        }

        private static bool Cagir(object o, string ad, params object[] args)
        {
            MethodInfo mi = o.GetType().GetMethods(Hepsi)
                .FirstOrDefault(m => m.Name == ad && m.GetParameters().Length == args.Length);
            if (mi == null) return false;
            mi.Invoke(mi.IsStatic ? null : o, args);
            return true;
        }
    }

    // =================================================================
    //  Çizili kesitleri bulma (ölçü ve çizgi tipi için ortak)
    // =================================================================
    internal class KesitCizimi
    {
        public string Ad;            // "SECTION A" -> "A"
        public double U, V, W;       // etiket konumu (düzlem koordinatları)
        public bool XZ;              // true: Y=W düzleminde (OutlineDrawing'in çizdiği hali), false: Z=W (2D Aktar sonrası)
        public List<ObjectId> Ids = new List<ObjectId>();
        public List<double[]> Segs = new List<double[]>();   // { u1, v1, u2, v2 }
    }

    internal static class KesitCizimleri
    {
        private static readonly Regex KesitYazisi = new Regex(@"^SECTION\s+(\S+)$", RegexOptions.IgnoreCase);
        private const double DerinlikTol = 5000;

        /// <summary>Model alanındaki "SECTION x" yazılarını ve her birine ait çizgileri bulur.
        /// silinecekKatman verilirse o katmandaki nesneler silinir (önceki otomatik ölçüler).</summary>
        internal static List<KesitCizimi> Topla(Transaction tr, BlockTableRecord ms, string silinecekKatman)
        {
            var etiketler = new List<KesitCizimi>();
            var cizgiler = new List<KeyValuePair<ObjectId, Point3d[]>>();
            foreach (ObjectId id in ms)
            {
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent == null) continue;
                if (silinecekKatman != null && string.Equals(ent.Layer, silinecekKatman, StringComparison.OrdinalIgnoreCase))
                {
                    try { ent.UpgradeOpen(); ent.Erase(); } catch { /* kilitli katman vb. */ }
                    continue;
                }
                var t = ent as DBText;
                if (t != null)
                {
                    if (t.TextString == null || t.Height < 250) continue;
                    Match mt = KesitYazisi.Match(t.TextString.Trim());
                    if (!mt.Success) continue;
                    Vector3d n = t.Normal;
                    Point3d p = t.HorizontalMode == TextHorizontalMode.TextLeft && t.VerticalMode == TextVerticalMode.TextBase
                        ? t.Position : t.AlignmentPoint;
                    if (Math.Abs(n.Y) > 0.9) etiketler.Add(new KesitCizimi { Ad = mt.Groups[1].Value, U = p.X, V = p.Z, W = p.Y, XZ = true });
                    else if (Math.Abs(n.Z) > 0.9) etiketler.Add(new KesitCizimi { Ad = mt.Groups[1].Value, U = p.X, V = p.Y, W = p.Z, XZ = false });
                    continue;
                }
                var l = ent as Line;
                if (l != null && !l.Layer.StartsWith("SECTION_MARK", StringComparison.OrdinalIgnoreCase))
                    cizgiler.Add(new KeyValuePair<ObjectId, Point3d[]>(id, new[] { l.StartPoint, l.EndPoint }));
            }

            foreach (KesitCizimi e in etiketler)
            {
                // Bu etikete ait çizgiler: aynı düzlemde, en yakın etiket bu olan ve 40 m içinde kalanlar
                foreach (var kv in cizgiler)
                {
                    Point3d[] c = kv.Value;
                    double ua = c[0].X, ub = c[1].X;
                    double va = e.XZ ? c[0].Z : c[0].Y, vb = e.XZ ? c[1].Z : c[1].Y;
                    double wa = e.XZ ? c[0].Y : c[0].Z, wb = e.XZ ? c[1].Y : c[1].Z;
                    // Eğik kesitte OutlineDrawing düzleme 100 mm'ye kadar yakın elemanları da çizer; bunlar
                    // çizim düzleminin önünde/arkasında kalır. Kesitler birbirinden u ile ayrıldığı için derinlik toleransı geniş.
                    if (Math.Abs(wa - e.W) > DerinlikTol || Math.Abs(wb - e.W) > DerinlikTol) continue;
                    double um = (ua + ub) / 2, vm = (va + vb) / 2;
                    if (Math.Abs(um - e.U) > 15000 || Math.Abs(vm - e.V) > 40000) continue;
                    bool baskasi = etiketler.Any(o => o != e && o.XZ == e.XZ && Math.Abs(o.W - e.W) <= 1 &&
                                                      Math.Abs(um - o.U) < Math.Abs(um - e.U) && Math.Abs(vm - o.V) <= 40000);
                    if (baskasi) continue;
                    e.Ids.Add(kv.Key);
                    e.Segs.Add(new[] { ua, va, ub, vb });
                }
            }
            return etiketler;
        }
    }

    // =================================================================
    //  Kesit ölçüleri
    // =================================================================
    internal static class KesitOlculeri
    {
        internal const string Katman = "OTO_KESIT_OLCU";

        /// <summary>Model alanındaki her "SECTION x" çizimine ölçü koyar; önceki otomatik ölçüleri siler.</summary>
        internal static string Ekle()
        {
            Document doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return "Ölçü: açık çizim yok.";
            Database db = doc.Database;
            int adet = 0, kesit = 0, gorunus = 0;

            using (doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var ms = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);
                KatmaniHazirla(db, tr);
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                var olcuKatmani = (LayerTableRecord)tr.GetObject(lt[Katman], OpenMode.ForRead);
                if (olcuKatmani.IsLocked) { olcuKatmani.UpgradeOpen(); olcuKatmani.IsLocked = false; }

                foreach (KesitCizimi e in KesitCizimleri.Topla(tr, ms, Katman))
                {
                    if (e.Segs.Count == 0) continue;
                    List<OlcuTanimi> olculer = OlcuPlani.Hesapla(e.Segs);
                    Matrix3d m = e.XZ
                        ? Matrix3d.Displacement(new Vector3d(0, e.W, 0)) * Matrix3d.Rotation(Math.PI / 2, Vector3d.XAxis, Point3d.Origin)
                        : Matrix3d.Displacement(new Vector3d(0, 0, e.W));
                    foreach (OlcuTanimi o in olculer) { OlcuKoy(db, tr, ms, o, m); adet++; }
                    kesit++;
                }

                // ---- ön (TRANSVERSE) ve yan (LONGITUDINAL) görünüş ölçüleri
                try
                {
                    Form f = FormBaglayici.AcikFormuBul();
                    string yol = f != null ? FormBaglayici.YolAl(f) : "";
                    KuleVerisi kule = yol.Length > 0 && File.Exists(yol) ? KuleOkuyucu.Oku(f.GetType().Assembly, yol) : null;
                    if (kule != null)
                    {
                        var cizgiler = new List<Point3d[]>();
                        foreach (ObjectId id in ms)
                        {
                            var l = tr.GetObject(id, OpenMode.ForRead) as Line;
                            if (l != null) cizgiler.Add(new[] { l.StartPoint, l.EndPoint });
                        }
                        foreach (var yuz in new[] { kule.OnYuz, kule.YanYuz })
                        {
                            if (yuz.Count == 0) continue;
                            int ofs; bool ikiD;
                            if (!GorunusBul(yuz, cizgiler, out ofs, out ikiD)) continue;
                            // Ölçüler görünüş başlığının düzleminde: çizildiği hal Y=0 (normal -Y), 2D Aktar sonrası Z=0
                            Matrix3d m = ikiD ? Matrix3d.Identity : Matrix3d.Rotation(Math.PI / 2, Vector3d.XAxis, Point3d.Origin);
                            foreach (OlcuTanimi o in GorunusOlcuPlani.Hesapla(yuz, ofs)) { OlcuKoy(db, tr, ms, o, m); adet++; }
                            gorunus++;
                        }
                    }
                }
                catch (System.Exception ex) { Komutlar.Yaz("\nGörünüş ölçüsü konamadı: " + ex.Message + "\n"); }
                tr.Commit();
            }
            string ozet = string.Format("Ölçü: {0} kesit, {1} görünüş", kesit, gorunus);
            Komutlar.Yaz(string.Format("\nÖlçü ayrıntı: {0} kesit + {1} görünüşe {2} ölçü (katman {3}).\n", kesit, gorunus, adet, Katman));
            Komutlar.Yaz("\n" + ozet + ".\n");
            return ozet;
        }

        private static void OlcuKoy(Database db, Transaction tr, BlockTableRecord ms, OlcuTanimi o, Matrix3d m)
        {
            var dim = new RotatedDimension(o.Rotation,
                new Point3d(o.U1, o.V1, 0), new Point3d(o.U2, o.V2, 0), new Point3d(o.UD, o.VD, 0),
                "", db.Dimstyle);
            dim.SetDatabaseDefaults(db);
            dim.Layer = Katman;
            dim.Dimtxt = o.Yazi;
            dim.Dimasz = 0.75 * o.Yazi;
            dim.Dimexo = 0.5 * o.Yazi;
            dim.Dimexe = 0.5 * o.Yazi;
            dim.Dimgap = 0.25 * o.Yazi;
            dim.Dimdec = 0;
            dim.Dimtad = 1;
            dim.Dimtih = false;
            dim.Dimtoh = false;
            dim.DimfxlenOn = true;   // uzatma çizgisi sabit boy: geometriden uzaktaki köşelerde
            dim.Dimfxlen = o.Uzatma; // çizim boyunca uzanan uzun çizgi olmasın
            dim.Dimscale = 1;     // çizimin ölçü stili ölçekli olsa da yazı boyu korunsun
            dim.Dimlfac = 1;      // ölçülen değer gerçek mm olsun
            dim.Dimtfac = 1;
            dim.Dimlunit = 2;     // ondalık
            dim.Dimrnd = 0;
            dim.TransformBy(m);
            ms.AppendEntity(dim);
            tr.AddNewlyCreatedDBObject(dim, true);
        }

        /// <summary>Görünüşün çizimde hangi ofsetle (0 / 50000) ve hangi halde (çizildiği gibi / 2D Aktar sonrası)
        /// durduğunu, çizgilerinin en az yarısı eşleşen seçenekle bulur.</summary>
        private static bool GorunusBul(List<YuzCizgisi> yuz, List<Point3d[]> cizgiler, out int ofs, out bool ikiD)
        {
            ofs = 0; ikiD = false;
            int enIyi = 0;
            foreach (int o in YuzEslestirici.Ofsetler)
                foreach (bool iki in new[] { false, true })
                {
                    var es = new YuzEslestirici(yuz, new[] { o }, !iki, iki);
                    string g;
                    int n = cizgiler.Count(c => es.Bul(c[0].X, c[0].Y, c[0].Z, c[1].X, c[1].Y, c[1].Z, out g));
                    if (n > enIyi) { enIyi = n; ofs = o; ikiD = iki; }
                }
            return enIyi >= Math.Max(1, yuz.Count / 2);
        }

        private static void KatmaniHazirla(Database db, Transaction tr)
        {
            var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (lt.Has(Katman)) return;
            lt.UpgradeOpen();
            var ltr = new LayerTableRecord { Name = Katman };
            lt.Add(ltr);
            tr.AddNewlyCreatedDBObject(ltr, true);
        }
    }

    // =================================================================
    //  Kesitlerde ana / redundant çizgi tipi
    // =================================================================
    internal static class KesitStili
    {
        public static short AnaRenk = 4;               // cyan, düz çizgi
        public static short RedundantRenk = 5;         // mavi, kesikli
        public static short MetinRenk = 7;             // eleman isimleri tek renk (beyaz/siyah)
        public const string CizgiTipiAdi = "OTO_KESIK";
        public static double Cizgi = 150, Bosluk = 75; // kesik deseni (mm, ekranda)

        private const BindingFlags Hepsi = BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;
        private static readonly string[] NoktaSutunlari = { "ColumnFirstPoint", "ColumnSecondPoint", "ColumnThirdPoint", "ColumnFourthPoint" };

        /// <returns>Kullanıcıya gösterilecek kısa özet</returns>
        internal static string Uygula(Form f)
        {
            Document doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return "Redundant stili: açık çizim yok.";
            Database db = doc.Database;

            var txt = FormBaglayici.Alan(f, "txtTowerPath") as TextBox;
            string path = txt != null ? (txt.Text ?? "").Trim() : "";
            if (path.Length == 0 || !File.Exists(path)) return "Redundant stili: .tow dosyası seçili değil.";
            KuleVerisi kule = KuleOkuyucu.Oku(f.GetType().Assembly, path);
            if (kule == null) return "Redundant stili: .tow okunamadı.";
            if (kule.RedundantGruplar.Count == 0) return "Redundant stili: .tow'da redundant grup yok (grup tipi Redundant veya açıklamada 'Redundant').";

            Dictionary<string, KesitTanimi> tanimlar = GridTanimlari(f);
            int ana = 0, red = 0, eslesmeyen = 0, kesit = 0, hata = 0, yuzAna = 0, yuzRed = 0, yazi = 0;
            var tanimsiz = new List<string>();
            string ilkHata = null;

            using (doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var ms = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);
                ObjectId kesik = KesikCizgiTipi(db, tr);
                ObjectId duz = SymbolUtilityServices.GetLinetypeContinuousId(db);
                double olcek = DesenOlcegi(db);

                // ---- ön / yan görünüş ve 3D model (YÜKLE): çizgiler koordinatından tanınır
                var model = kule.Uyeler.Values.Select(m => new YuzCizgisi
                {
                    P = new[] { m.start_x, m.start_y, m.start_z, m.end_x, m.end_y, m.end_z },
                    Grup = m.group_label
                });
                var yuz = new YuzEslestirici(kule.OnYuz.Concat(kule.YanYuz).Concat(model));
                foreach (ObjectId id in ms)
                {
                    var l = tr.GetObject(id, OpenMode.ForRead) as Line;
                    if (l == null) continue;
                    Point3d a = l.StartPoint, b = l.EndPoint;
                    string grup;
                    if (!yuz.Bul(a.X, a.Y, a.Z, b.X, b.Y, b.Z, out grup)) continue;
                    try
                    {
                        l.UpgradeOpen();
                        if (Boya(l, grup != null && kule.RedundantGruplar.Contains(grup.Trim()), kesik, duz, olcek)) yuzRed++; else yuzAna++;
                    }
                    catch (System.Exception ex) { hata++; if (ilkHata == null) ilkHata = ex.Message; }
                }

                foreach (KesitCizimi e in KesitCizimleri.Topla(tr, ms, null))
                {
                    if (e.Segs.Count == 0) continue;
                    KesitTanimi tanim;
                    if (!tanimlar.TryGetValue(e.Ad, out tanim)) { tanimsiz.Add(e.Ad); continue; }
                    List<string> anahtarlar;
                    List<double[]> kaynak = KesitSecici.Sec(kule.Uyeler, tanim, out anahtarlar);
                    int[] eslesme = Eslestirici.Eslestir(kaynak, e.Segs, 2.0);
                    kesit++;
                    for (int i = 0; i < e.Ids.Count; i++)
                    {
                        try
                        {
                            var ent = (Entity)tr.GetObject(e.Ids[i], OpenMode.ForWrite);
                            if (eslesme[i] < 0)
                            {
                                // Tanınamayan çizgi: tek renk olsun diye ana eleman kabul edilir
                                Boya(ent, false, kesik, duz, olcek);
                                eslesmeyen++;
                                continue;
                            }
                            FinalMember m = kule.Uyeler[anahtarlar[eslesme[i]]];
                            bool redundant = m.group_label != null && kule.RedundantGruplar.Contains(m.group_label.Trim());
                            if (Boya(ent, redundant, kesik, duz, olcek)) red++; else ana++;
                        }
                        catch (System.Exception ex)
                        {
                            hata++;
                            if (ilkHata == null) ilkHata = ex.Message;
                        }
                    }
                }
                // ---- eleman isimleri: OutlineDrawing katman rengini (section label) alıyordu; hepsi tek renk
                foreach (ObjectId id in ms)
                {
                    var t = tr.GetObject(id, OpenMode.ForRead) as DBText;
                    if (t == null || t.Height >= 250) continue;
                    string kat = t.Layer ?? "";
                    if (!kat.StartsWith("SECTION ", StringComparison.OrdinalIgnoreCase) && kat != "0") continue;
                    try
                    {
                        t.UpgradeOpen();
                        t.Color = AcColor.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, MetinRenk);
                        yazi++;
                    }
                    catch (System.Exception ex) { hata++; if (ilkHata == null) ilkHata = ex.Message; }
                }
                tr.Commit();
            }
            try { doc.Editor.Regen(); } catch { /* görüntü yenilenemezse REGEN elle */ }

            string ozet = string.Format("Çizgi: {0} kesikli, {1} düz", yuzRed + red, yuzAna + ana);
            if (eslesmeyen > 0) ozet += ", tanınmayan " + eslesmeyen;
            if (tanimsiz.Count > 0) ozet += ", tabloda olmayan kesit: " + string.Join(",", tanimsiz);
            if (hata > 0) ozet += string.Format(", {0} hata ({1})", hata, ilkHata);
            Komutlar.Yaz(string.Format("\nÇizgi tipi ayrıntı: görünüşler {0} redundant / {1} ana; {2} kesit {3} redundant / {4} ana.\n",
                yuzRed, yuzAna, kesit, red, ana));
            Komutlar.Yaz("\n" + ozet + ".\n");
            return ozet;
        }

        /// <summary>Kesik desenin ekranda her zaman 150/75 mm görünmesi için çizgi tipi ölçeği. Model alanında görünen desen
        /// = LTSCALE × nesne ölçeği × (MSLTSCALE açıksa anotasyon ölçeğinin çarpanı, ör. 1:100 → 100). Bu iki ayar
        /// büyükse 150 mm'lik desen çizgi boyundan uzun olur ve AutoCAD çizgiyi düz gösterir.</summary>
        private static double DesenOlcegi(Database db)
        {
            double lts = db.Ltscale > 1e-9 ? db.Ltscale : 1.0, anno = 1.0;
            try
            {
                if (Convert.ToInt32(AcApp.GetSystemVariable("MSLTSCALE")) != 0 && db.Cannoscale != null && db.Cannoscale.Scale > 1e-9) anno = 1.0 / db.Cannoscale.Scale;
            }
            catch { /* eski sürüm: anotasyon ölçeği yok */ }
            Komutlar.Yaz(string.Format(CultureInfo.InvariantCulture,
                "\nKesik çizgi: LTSCALE {0:0.###}, anotasyon çarpanı {1:0.###} -> nesne çizgi tipi ölçeği {2:0.#####}.\n", lts, anno, 1.0 / (lts * anno)));
            return 1.0 / (lts * anno);
        }

        /// <returns>redundant ise true</returns>
        private static bool Boya(Entity ent, bool redundant, ObjectId kesik, ObjectId duz, double olcek)
        {
            if (redundant)
            {
                ent.Color = AcColor.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, RedundantRenk);
                ent.LinetypeId = kesik;
                ent.LinetypeScale = olcek;
            }
            else
            {
                ent.Color = AcColor.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, AnaRenk);
                ent.LinetypeId = duz;
            }
            return redundant;
        }

        /// <summary>Dosyaya (acad.lin) bağımlı olmadan kesikli çizgi tipini oluşturur.</summary>
        private static ObjectId KesikCizgiTipi(Database db, Transaction tr)
        {
            var ltt = (LinetypeTable)tr.GetObject(db.LinetypeTableId, OpenMode.ForRead);
            if (ltt.Has(CizgiTipiAdi)) return ltt[CizgiTipiAdi];
            ltt.UpgradeOpen();
            var r = new LinetypeTableRecord
            {
                Name = CizgiTipiAdi,
                AsciiDescription = "Oto Kesit redundant __ __ __",
                PatternLength = Cizgi + Bosluk,
                NumDashes = 2
            };
            r.SetDashLengthAt(0, Cizgi);
            r.SetDashLengthAt(1, -Bosluk);
            ObjectId id = ltt.Add(r);
            tr.AddNewlyCreatedDBObject(r, true);
            return id;
        }

        /// <summary>Kesit tablosundaki satırları doğrudan okur (Kesit İsmi, Yükseklik, 1-4. Nokta "x;y;z").</summary>
        private static Dictionary<string, KesitTanimi> GridTanimlari(Form f)
        {
            var sonuc = new Dictionary<string, KesitTanimi>(StringComparer.OrdinalIgnoreCase);
            var grid = FormBaglayici.Alan(f, "dataGridView1") as DataGridView;
            if (grid == null) return sonuc;
            foreach (DataGridViewRow row in grid.Rows)
            {
                if (row.IsNewRow || !grid.Columns.Contains("ColumnSection")) continue;
                string ad = Hucre(row, "ColumnSection");
                if (ad.Length == 0) continue;
                var tanim = new KesitTanimi();
                double h;
                string hs = Hucre(row, "ColumnHeight");
                if (hs.Length > 0 && Sayi(hs, out h)) tanim.Yukseklik = h;
                foreach (string sutun in NoktaSutunlari)
                {
                    string[] p = Hucre(row, sutun).Split(';');
                    double x, y, z;
                    if (p.Length == 3 && Sayi(p[0], out x) && Sayi(p[1], out y) && Sayi(p[2], out z))
                        tanim.Noktalar.Add(new[] { x, y, z });
                }
                if (tanim.Yukseklik.HasValue || tanim.Noktalar.Count >= 3) sonuc[ad] = tanim;
            }
            return sonuc;
        }

        private static string Hucre(DataGridViewRow row, string sutun)
        {
            if (!row.DataGridView.Columns.Contains(sutun)) return "";
            object v = row.Cells[sutun].Value;
            return v == null ? "" : v.ToString().Trim();
        }

        // OutlineDrawing.TryParseDouble ile aynı: önce geçerli kültür, sonra Invariant
        private static bool Sayi(string s, out double d)
        {
            return double.TryParse(s, NumberStyles.Float, CultureInfo.CurrentCulture, out d) ||
                   double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d);
        }
    }

    /// <summary>Grid'deki bir kesitin tanımı: yükseklik veya düzlem noktaları.</summary>
    public class KesitTanimi
    {
        public double? Yukseklik;
        public List<double[]> Noktalar = new List<double[]>();
    }

    /// <summary>OutlineDrawing'in SectionDetection seçim kuralının kopyası; seçilen elemanların
    /// düzlemdeki 2B izdüşümünü verir (çizimle eşleştirmek için, döndürme/aynalama önemsiz).</summary>
    public static class KesitSecici
    {
        public static double YukseklikTol = 15;     // runtowerfile -> SectionDetection(tol = 15)
        public static double DuzlemTol = 100;       // SectionDetection: orta nokta düzleme < 100 mm

        public static List<double[]> Sec(IDictionary<string, FinalMember> uyeler, KesitTanimi t, out List<string> anahtarlar)
        {
            var segs = new List<double[]>();
            anahtarlar = new List<string>();
            if (t.Yukseklik.HasValue)
            {
                double h = t.Yukseklik.Value;
                foreach (var kv in uyeler)
                {
                    var m = kv.Value;
                    if (Math.Abs(m.start_z - h) > YukseklikTol || Math.Abs(m.end_z - h) > YukseklikTol) continue;
                    segs.Add(new[] { m.start_x, m.start_y, m.end_x, m.end_y });
                    anahtarlar.Add(kv.Key);
                }
                return segs;
            }
            if (t.Noktalar.Count < 3) return segs;

            double[] p0 = t.Noktalar[0], p1 = t.Noktalar[1], pn = t.Noktalar[t.Noktalar.Count - 1];
            double[] n = Norm(Cross(Sub(p1, p0), Sub(pn, p0)));
            if (n == null) return segs;
            double[] e1 = Norm(Sub(p1, p0));
            double[] e2 = Cross(n, e1);
            Func<double[], double[]> proj = q => { var d = Sub(q, p0); return new[] { Dot(d, e1), Dot(d, e2) }; };

            var poly = t.Noktalar.Select(proj).ToList();
            double cx = poly.Average(q => q[0]), cy = poly.Average(q => q[1]);
            poly = poly.OrderBy(q => Math.Atan2(q[1] - cy, q[0] - cx)).ToList();

            foreach (var kv in uyeler)
            {
                var m = kv.Value;
                double[] a = { m.start_x, m.start_y, m.start_z }, b = { m.end_x, m.end_y, m.end_z };
                double[] mid = { (a[0] + b[0]) / 2, (a[1] + b[1]) / 2, (a[2] + b[2]) / 2 };
                if (Math.Abs(Dot(Sub(mid, p0), n)) >= DuzlemTol) continue;
                double[] pa = proj(a), pb = proj(b), pm = proj(mid);
                if (!IcindeMi(pm, poly) && !IcindeMi(pa, poly) && !IcindeMi(pb, poly)) continue;
                segs.Add(new[] { pa[0], pa[1], pb[0], pb[1] });
                anahtarlar.Add(kv.Key);
            }
            return segs;
        }

        // OutlineDrawing.PointInPolygon ile aynı (ışın, +1.0 tolerans)
        private static bool IcindeMi(double[] p, List<double[]> poly)
        {
            bool ic = false;
            for (int i = 0, j = poly.Count - 1; i < poly.Count; j = i++)
                if ((poly[i][1] > p[1]) != (poly[j][1] > p[1]) &&
                    p[0] < (poly[j][0] - poly[i][0]) * (p[1] - poly[i][1]) / (poly[j][1] - poly[i][1]) + poly[i][0] + 1.0)
                    ic = !ic;
            return ic;
        }

        private static double[] Sub(double[] a, double[] b) { return new[] { a[0] - b[0], a[1] - b[1], a[2] - b[2] }; }
        private static double Dot(double[] a, double[] b) { return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]; }
        private static double[] Cross(double[] a, double[] b) { return new[] { a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2], a[0] * b[1] - a[1] * b[0] }; }
        private static double[] Norm(double[] a) { double l = Math.Sqrt(Dot(a, a)); return l < 1e-9 ? null : new[] { a[0] / l, a[1] / l, a[2] / l }; }
    }

    /// <summary>İki 2B çizgi kümesini rijit dönüşümle (döndürme + öteleme, gerekirse aynalama) eşleştirir.</summary>
    public static class Eslestirici
    {
        /// <returns>hedef[i] için eşleşen kaynak indeksi, yoksa -1</returns>
        public static int[] Eslestir(List<double[]> kaynak, List<double[]> hedef, double tol)
        {
            var sonuc = Enumerable.Repeat(-1, hedef.Count).ToArray();
            if (kaynak.Count == 0 || hedef.Count == 0) return sonuc;

            Func<double[], double> boy = s => Math.Sqrt((s[2] - s[0]) * (s[2] - s[0]) + (s[3] - s[1]) * (s[3] - s[1]));
            var capalar = Enumerable.Range(0, kaynak.Count).OrderByDescending(i => boy(kaynak[i])).Take(3).ToList();

            int enIyi = -1; int[] enIyiEslesme = null; List<double[]> enIyiDonmus = null;
            foreach (int ci in capalar)
            {
                double[] s = kaynak[ci];
                double ls = boy(s);
                for (int hi = 0; hi < hedef.Count; hi++)
                {
                    double[] d = hedef[hi];
                    if (Math.Abs(boy(d) - ls) > tol) continue;
                    foreach (bool ayna in new[] { false, true })
                        foreach (bool ters in new[] { false, true })
                        {
                            double fy = ayna ? -1 : 1;
                            double ax = s[0], ay = fy * s[1], bx = s[2], by = fy * s[3];
                            double cx = ters ? d[2] : d[0], cy = ters ? d[3] : d[1], dx = ters ? d[0] : d[2], dy = ters ? d[1] : d[3];
                            double th = Math.Atan2(dy - cy, dx - cx) - Math.Atan2(by - ay, bx - ax);
                            double co = Math.Cos(th), si = Math.Sin(th);
                            double tx = cx - (co * ax - si * ay), ty = cy - (si * ax + co * ay);
                            Func<double, double, double[]> T = (x, y) => { y *= fy; return new[] { co * x - si * y + tx, si * x + co * y + ty }; };

                            var donmus = kaynak.Select(k => { var p = T(k[0], k[1]); var q = T(k[2], k[3]); return new[] { p[0], p[1], q[0], q[1] }; }).ToList();
                            int[] es = Esle(donmus, hedef, tol);
                            int skor = es.Count(x => x >= 0);
                            if (skor > enIyi) { enIyi = skor; enIyiEslesme = es; enIyiDonmus = donmus; }
                            if (enIyi == hedef.Count) return enIyiEslesme;
                        }
                }
            }
            if (enIyiEslesme == null) return sonuc;

            // İkinci geçiş: bulunan en iyi dönüşümle eşleşmeyen çizgiler için toleranslı (GenisTol) en yakın eleman
            for (int h = 0; h < hedef.Count; h++)
            {
                if (enIyiEslesme[h] >= 0) continue;
                double[] d = hedef[h];
                double enAz = double.MaxValue; int secilen = -1;
                for (int i = 0; i < enIyiDonmus.Count; i++)
                {
                    double[] k = enIyiDonmus[i];
                    double d1 = Math.Max(Uz(k[0], k[1], d[0], d[1]), Uz(k[2], k[3], d[2], d[3]));
                    double d2 = Math.Max(Uz(k[0], k[1], d[2], d[3]), Uz(k[2], k[3], d[0], d[1]));
                    double m = Math.Min(d1, d2);
                    if (m < enAz) { enAz = m; secilen = i; }
                }
                if (enAz <= GenisTol) enIyiEslesme[h] = secilen;
            }
            return enIyiEslesme;
        }

        public static double GenisTol = 60;   // mm, ikinci geçiş

        private static double Uz(double x1, double y1, double x2, double y2)
        {
            return Math.Sqrt((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2));
        }

        private static int[] Esle(List<double[]> kaynak, List<double[]> hedef, double tol)
        {
            var idx = new Dictionary<string, List<int>>();
            Func<double, double, string> k = (x, y) => Math.Round(x / 10).ToString(CultureInfo.InvariantCulture) + "," + Math.Round(y / 10).ToString(CultureInfo.InvariantCulture);
            for (int i = 0; i < kaynak.Count; i++)
            {
                string key = k((kaynak[i][0] + kaynak[i][2]) / 2, (kaynak[i][1] + kaynak[i][3]) / 2);
                List<int> l;
                if (!idx.TryGetValue(key, out l)) { l = new List<int>(); idx[key] = l; }
                l.Add(i);
            }
            var sonuc = new int[hedef.Count];
            for (int h = 0; h < hedef.Count; h++)
            {
                sonuc[h] = -1;
                double[] d = hedef[h];
                double mx = (d[0] + d[2]) / 2, my = (d[1] + d[3]) / 2;
                double gx = Math.Round(mx / 10), gy = Math.Round(my / 10);
                for (int ox = -1; ox <= 1 && sonuc[h] < 0; ox++)
                    for (int oy = -1; oy <= 1 && sonuc[h] < 0; oy++)
                    {
                        List<int> l;
                        if (!idx.TryGetValue((gx + ox).ToString(CultureInfo.InvariantCulture) + "," + (gy + oy).ToString(CultureInfo.InvariantCulture), out l)) continue;
                        foreach (int i in l)
                        {
                            double[] s = kaynak[i];
                            bool ayni = (Yakin(s[0], s[1], d[0], d[1], tol) && Yakin(s[2], s[3], d[2], d[3], tol)) ||
                                        (Yakin(s[0], s[1], d[2], d[3], tol) && Yakin(s[2], s[3], d[0], d[1], tol));
                            if (ayni) { sonuc[h] = i; break; }
                        }
                    }
            }
            return sonuc;
        }

        private static bool Yakin(double x1, double y1, double x2, double y2, double tol)
        {
            return Math.Abs(x1 - x2) <= tol && Math.Abs(y1 - y2) <= tol;
        }
    }

    /// <summary>Tek bir ölçünün düzlem (u, v) koordinatlarında tanımı. Rotation: 0 = yatay, π/2 = düşey.</summary>
    public class OlcuTanimi
    {
        public double U1, V1, U2, V2, UD, VD, Rotation;
        /// <summary>Kesit büyüklüğüne göre ölçü yazı yüksekliği ve uzatma çizgisi boyu (mm).</summary>
        public double Yazi = 200, Uzatma = 400;
        public double Deger { get { return Math.Abs(Rotation) < 1e-9 ? Math.Abs(U2 - U1) : Math.Abs(V2 - V1); } }
    }

    /// <summary>AutoCAD'den bağımsız ölçü yerleşimi: kesit çizgilerinin dış konturundan
    /// üstte zincir ölçü (kontur köşelerinin u değerleri), sağda toplam derinlik.</summary>
    public static class OlcuPlani
    {
        public static double Aralik = 500;      // ölçü çizgisinin kesitten uzaklığı (mm)
        public static double EnKisa = 50;       // bundan kısa zincir parçaları birleştirilir (mm)
        public static double KirikAcisi = 8;    // konturda bundan küçük dönüşler köşe sayılmaz (derece)

        /// <param name="segs">her eleman: { u1, v1, u2, v2 }</param>
        public static List<OlcuTanimi> Hesapla(List<double[]> segs)
        {
            var sonuc = new List<OlcuTanimi>();
            segs = AnaParcalar(segs);
            var pts = new List<double[]>();
            foreach (var s in segs) { pts.Add(new[] { s[0], s[1] }); pts.Add(new[] { s[2], s[3] }); }
            var hull = Sadelestir(Kontur(pts), KirikAcisi);
            if (hull.Count < 2) return sonuc;

            double minU = hull.Min(p => p[0]), maxU = hull.Max(p => p[0]);
            double minV = hull.Min(p => p[1]), maxV = hull.Max(p => p[1]);

            // Küçük kesitlerde (ör. 829 mm'lik tepe kapağı) ölçü aralığı ve yazı orantılı küçülür
            double boyut = Math.Max(maxU - minU, maxV - minV);
            double aralik = Math.Min(Aralik, Math.Max(150, 0.1 * boyut));
            double yazi = Math.Min(200, Math.Max(100, 0.06 * boyut));
            double uzatma = 0.8 * aralik;

            // Üst zincir: kontur köşelerinin u değerleri (her u için en üstteki köşe)
            var us = new List<double[]>();
            foreach (var p in hull.OrderBy(p => p[0]))
            {
                if (us.Count > 0 && p[0] - us[us.Count - 1][0] < EnKisa)
                {
                    if (p[1] > us[us.Count - 1][1]) us[us.Count - 1] = p;
                    continue;
                }
                us.Add(p);
            }
            if (us.Count >= 2 && maxU - minU >= EnKisa)
            {
                double vd = maxV + aralik;
                for (int i = 0; i + 1 < us.Count; i++)
                    sonuc.Add(new OlcuTanimi
                    {
                        U1 = us[i][0], V1 = us[i][1], U2 = us[i + 1][0], V2 = us[i + 1][1],
                        UD = (us[i][0] + us[i + 1][0]) / 2, VD = vd, Rotation = 0, Yazi = yazi, Uzatma = uzatma
                    });
            }

            // Sağ: toplam derinlik (en üst ve en alt köşelerden sağdakiler)
            if (maxV - minV >= EnKisa)
            {
                var ust = hull.Where(p => p[1] >= maxV - 1e-6).OrderByDescending(p => p[0]).First();
                var alt = hull.Where(p => p[1] <= minV + 1e-6).OrderByDescending(p => p[0]).First();
                sonuc.Add(new OlcuTanimi
                {
                    U1 = ust[0], V1 = ust[1], U2 = alt[0], V2 = alt[1],
                    UD = maxU + aralik, VD = (maxV + minV) / 2, Rotation = Math.PI / 2, Yazi = yazi, Uzatma = uzatma
                });
            }
            return sonuc;
        }

        public static double KucukParcaOrani = 0.1;   // toplam boyu en büyük parçanın bundan azı olan kopuk parçalar ölçülmez

        /// <summary>Birbirine bağlı çizgi grupları; toplam boyu en büyük grubun KucukParcaOrani'ndan az olan kopuk
        /// küçük gruplar (ör. kol ucundaki tek kısa eleman) ölçüye katılmaz, zincir ölçü boşluğa uzamaz.</summary>
        public static List<double[]> AnaParcalar(List<double[]> segs)
        {
            if (segs.Count < 2) return segs;
            Func<double, double, string> k = (x, y) => Math.Round(x).ToString(CultureInfo.InvariantCulture) + "," + Math.Round(y).ToString(CultureInfo.InvariantCulture);
            var ebeveyn = new Dictionary<string, string>();
            Func<string, string> bul = null;
            bul = a => { while (ebeveyn[a] != a) a = ebeveyn[a]; return a; };
            foreach (var s in segs)
            {
                string a = k(s[0], s[1]), b = k(s[2], s[3]);
                if (!ebeveyn.ContainsKey(a)) ebeveyn[a] = a;
                if (!ebeveyn.ContainsKey(b)) ebeveyn[b] = b;
                string ra = bul(a), rb = bul(b);
                if (ra != rb) ebeveyn[ra] = rb;
            }
            var gruplar = segs.GroupBy(s => bul(k(s[0], s[1]))).ToList();
            Func<IEnumerable<double[]>, double> boy = g => g.Sum(s => Math.Sqrt((s[2] - s[0]) * (s[2] - s[0]) + (s[3] - s[1]) * (s[3] - s[1])));
            double enBuyuk = gruplar.Max(g => boy(g));
            var kalan = gruplar.Where(g => boy(g) >= KucukParcaOrani * enBuyuk).SelectMany(g => g).ToList();
            return kalan.Count > 0 ? kalan : segs;
        }

        /// <summary>Konturda neredeyse düz devam eden (dönüşü açı sınırından küçük) köşeleri atar;
        /// böylece travers kolu üzerindeki küçük kırıklar zincir ölçüyü bölmez.</summary>
        private static List<double[]> Sadelestir(List<double[]> hull, double aciDerece)
        {
            var h = new List<double[]>(hull);
            bool degisti = true;
            while (degisti && h.Count > 3)
            {
                degisti = false;
                for (int i = 0; i < h.Count && h.Count > 3; i++)
                {
                    double[] a = h[(i - 1 + h.Count) % h.Count], b = h[i], c = h[(i + 1) % h.Count];
                    double a1 = Math.Atan2(b[1] - a[1], b[0] - a[0]), a2 = Math.Atan2(c[1] - b[1], c[0] - b[0]);
                    double donus = Math.Abs(a2 - a1) * 180 / Math.PI;
                    if (donus > 180) donus = 360 - donus;
                    if (donus < aciDerece) { h.RemoveAt(i); degisti = true; i--; }
                }
            }
            return h;
        }

        private static List<double[]> Kontur(List<double[]> pts)
        {
            var p = pts.OrderBy(q => q[0]).ThenBy(q => q[1]).ToList();
            var uniq = new List<double[]>();
            foreach (var q in p)
                if (uniq.Count == 0 || Math.Abs(q[0] - uniq[uniq.Count - 1][0]) > 1e-6 || Math.Abs(q[1] - uniq[uniq.Count - 1][1]) > 1e-6)
                    uniq.Add(q);
            if (uniq.Count < 3) return uniq;
            Func<double[], double[], double[], double> cr = (o, a, b) => (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0]);
            var h = new List<double[]>();
            foreach (var q in uniq)
            {
                while (h.Count >= 2 && cr(h[h.Count - 2], h[h.Count - 1], q) <= 1e-9) h.RemoveAt(h.Count - 1);
                h.Add(q);
            }
            int lower = h.Count + 1;
            for (int i = uniq.Count - 2; i >= 0; i--)
            {
                var q = uniq[i];
                while (h.Count >= lower && cr(h[h.Count - 2], h[h.Count - 1], q) <= 1e-9) h.RemoveAt(h.Count - 1);
                h.Add(q);
            }
            h.RemoveAt(h.Count - 1);
            return h;
        }
    }

    /// <summary>
    /// Ön (TRANSVERSE) / yan (LONGITUDINAL) görünüş ölçüleri, görünüş koordinatlarında (u = x + ofset, v = z):
    ///  - sağda düşey zincir: PLS-TOWER kısımlarının (Leg Ext., Body Ext., Basic Body, Crossarms, Pikes ...) sınır
    ///    kotları, onun dışında toplam yükseklik;
    ///  - altta (görünüş başlığının altında) taban genişliği;
    ///  - kol / tepe uçlarında eksenden uca yatay ölçü (en üstteki uçta üstte, diğerlerinde ucun altında).
    /// </summary>
    public static class GorunusOlcuPlani
    {
        public static double Yazi = 250, Aralik = 700;
        public static double BirlesmeTol = 50;      // bundan yakın kotlar tek kot sayılır (mm)
        public static double UcTol = 100;           // uç (kol, tepe) sayılması için komşu kotlardan taşma (mm)
        private const double BaslikBoyu = 600, BaslikPayi = 890;   // OutlineDrawing: başlık 600, taban - 890'da

        public static List<OlcuTanimi> Hesapla(List<YuzCizgisi> yuz, double ofs)
        {
            var sonuc = new List<OlcuTanimi>();
            var segs = yuz.Select(c => new[] { c.P[0] + ofs, c.P[2], c.P[3] + ofs, c.P[5] }).ToList();
            if (segs.Count == 0) return sonuc;
            double minU = segs.Min(s => Math.Min(s[0], s[2])), maxU = segs.Max(s => Math.Max(s[0], s[2]));
            double minV = segs.Min(s => Math.Min(s[1], s[3])), maxV = segs.Max(s => Math.Max(s[1], s[3]));
            if (maxV - minV < 1000) return sonuc;
            Func<OlcuTanimi, OlcuTanimi> t = o => { o.Yazi = Yazi; o.Uzatma = 0.8 * Aralik; return o; };

            // ---- düşey zincir: kısım sınırları
            var kotlar = new List<double> { minV, maxV };
            foreach (var g in yuz.Where(c => !string.IsNullOrWhiteSpace(c.Kesim) && c.Kesim != "UNDEFINED").GroupBy(c => c.Kesim.Trim()))
            {
                kotlar.Add(g.Min(c => Math.Min(c.P[2], c.P[5])));
                kotlar.Add(g.Max(c => Math.Max(c.P[2], c.P[5])));
            }
            kotlar = Birlestir(kotlar);
            double uz = maxU + Aralik;
            for (int i = 0; i + 1 < kotlar.Count; i++)
                sonuc.Add(t(new OlcuTanimi
                {
                    U1 = Sag(segs, kotlar[i]), V1 = kotlar[i], U2 = Sag(segs, kotlar[i + 1]), V2 = kotlar[i + 1],
                    UD = uz, VD = (kotlar[i] + kotlar[i + 1]) / 2, Rotation = Math.PI / 2
                }));
            if (kotlar.Count > 2)
                sonuc.Add(t(new OlcuTanimi
                {
                    U1 = Sag(segs, minV), V1 = minV, U2 = Sag(segs, maxV), V2 = maxV,
                    UD = uz + Aralik, VD = (minV + maxV) / 2, Rotation = Math.PI / 2
                }));

            // ---- taban genişliği (başlığın altında)
            double l0 = Sol(segs, minV), r0 = Sag(segs, minV), eksen = (l0 + r0) / 2;
            if (r0 - l0 > BirlesmeTol)
                sonuc.Add(t(new OlcuTanimi
                {
                    U1 = l0, V1 = minV, U2 = r0, V2 = minV,
                    UD = eksen, VD = minV - BaslikPayi - Aralik, Rotation = 0
                }));

            // ---- uçlar: sağ / sol profilin yerel en büyükleri (taban hariç)
            var seviyeler = Birlestir(segs.SelectMany(s => new[] { s[1], s[3] }).ToList(), 1);
            var sag = seviyeler.Select(v => Sag(segs, v) - eksen).ToList();
            var sol = seviyeler.Select(v => eksen - Sol(segs, v)).ToList();
            var ucSag = YerelEnBuyuk(sag); var ucSol = YerelEnBuyuk(sol);
            foreach (int i in ucSag.Union(ucSol).Distinct().OrderBy(i => i))
            {
                double v = seviyeler[i];
                if (v <= minV + BirlesmeTol) continue;
                double ur = eksen + sag[i], ul = eksen - sol[i];
                bool sagUc = ucSag.Contains(i), solUc = ucSol.Contains(i);
                double bas = solUc ? ul : eksen, son = sagUc ? ur : eksen;
                // Ucun üstünde bu aralıkta çizgi yoksa ölçü üste, varsa alta
                bool ustBos = !segs.Any(s => Math.Max(s[1], s[3]) > v + BirlesmeTol && Math.Max(s[0], s[2]) > bas && Math.Min(s[0], s[2]) < son);
                double vd = ustBos ? v + Aralik : v - Aralik;
                if (solUc) sonuc.Add(t(new OlcuTanimi { U1 = ul, V1 = v, U2 = eksen, V2 = v, UD = (ul + eksen) / 2, VD = vd, Rotation = 0 }));
                if (sagUc) sonuc.Add(t(new OlcuTanimi { U1 = eksen, V1 = v, U2 = ur, V2 = v, UD = (ur + eksen) / 2, VD = vd, Rotation = 0 }));
            }
            return sonuc;
        }

        private static List<int> YerelEnBuyuk(List<double> r)
        {
            var l = new List<int>();
            for (int i = 1; i < r.Count; i++)
            {
                // Düz kısımda (aynı genişlik) ilerleyen kotlar: platonun sonu alınır
                int j = i;
                while (j + 1 < r.Count && Math.Abs(r[j + 1] - r[i]) <= 1) j++;
                bool alt = r[i] > r[i - 1] + UcTol;
                bool ust = j + 1 >= r.Count || r[j + 1] < r[i] - UcTol;
                if (alt && ust) l.Add(i);
                i = j;
            }
            return l;
        }

        private static List<double> Birlestir(List<double> kotlar, double tol = -1)
        {
            if (tol < 0) tol = BirlesmeTol;
            var l = new List<double>();
            foreach (double v in kotlar.OrderBy(x => x))
                if (l.Count == 0 || v - l[l.Count - 1] > tol) l.Add(v);
            return l;
        }

        /// <summary>v kotunda görünüşün en sağ / en sol u değeri (kotu kesen veya o kotta biten çizgiler).</summary>
        private static double Sag(List<double[]> segs, double v) { return Uc(segs, v, true); }
        private static double Sol(List<double[]> segs, double v) { return Uc(segs, v, false); }

        private static double Uc(List<double[]> segs, double v, bool sag)
        {
            double sonuc = sag ? double.MinValue : double.MaxValue;
            foreach (var s in segs)
            {
                double v1 = Math.Min(s[1], s[3]), v2 = Math.Max(s[1], s[3]);
                if (v < v1 - 1 || v > v2 + 1) continue;
                double u;
                if (v2 - v1 < 1) u = sag ? Math.Max(s[0], s[2]) : Math.Min(s[0], s[2]);
                else u = s[0] + (s[2] - s[0]) * (v - s[1]) / (s[3] - s[1]);
                sonuc = sag ? Math.Max(sonuc, u) : Math.Min(sonuc, u);
            }
            return sonuc == double.MinValue || sonuc == double.MaxValue ? 0 : sonuc;
        }
    }

    /// <summary>Model alanında OutlineDrawing'e ait olmayan nesneler (başka katmanlarda).</summary>
    internal static class YabanciNesneler
    {
        internal static bool Bizim(string katman)
        {
            if (string.IsNullOrEmpty(katman)) return true;
            return katman == "0"
                || katman.StartsWith("SECTION", StringComparison.OrdinalIgnoreCase)
                || string.Equals(katman, KesitOlculeri.Katman, StringComparison.OrdinalIgnoreCase)
                || string.Equals(katman, "Defpoints", StringComparison.OrdinalIgnoreCase);
        }

        internal static int Say()
        {
            Document doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return 0;
            int n = 0;
            using (doc.LockDocument())
            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                var ms = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(doc.Database), OpenMode.ForRead);
                foreach (ObjectId id in ms)
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent == null || Bizim(ent.Layer) || ent is Table) continue;
                    var t = ent as DBText;
                    if (t != null && t.TextString != null && t.TextString.Trim().EndsWith("FACE", StringComparison.OrdinalIgnoreCase)) continue;
                    n++;
                }
                tr.Commit();
            }
            return n;
        }
    }

    /// <summary>
    /// OutlineDrawing'in FinalizeMember'ı yalnız 1 (X), 2 (Y), 3 (XY) simetri kodlarını çoğaltıyor; kod 12 olan
    /// elemanların aynaları çizilmiyordu (örnek kulede sol traverste 14 eleman eksik). Tek DLL derlemesinde
    /// FinalizeMember'ın sonuna (Z normalizasyonundan önce) bu metodun çağrısı eklenir (eklenti/tek_dll.sh).
    /// Kod 12 için yalnız X kopyası (PLS koordinatında y -> -y, diğer travers) üretilir; geometrisi zaten var
    /// olan atlanır. Gerekçe: aynı kulede C3-1 (kod 1, S->Y) ile C3-2 (kod 12, Y->S) aynı tip enine eleman; kod 12
    /// yön tersine çevrilmiş X simetrisi gibi davranıyor. Y aynası alınırsa zikzak çaprazlar X çaprazlamaya döner (yanlış).
    /// </summary>
    public static class SimetriDuzeltici
    {
        private const BindingFlags Hepsi = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
        public static readonly int[] DuzeltilenKodlar = { 12 };
        /// <summary>Son çalıştırmada eklenen eleman sayısı.</summary>
        public static int Eklenen;

        public static void Uygula(object dataprocessing)
        {
            try
            {
                Eklenen = 0;
                Type t = dataprocessing.GetType();
                var am = t.GetProperty("angMember", Hepsi).GetValue(dataprocessing, null) as IDictionary;
                var fm = t.GetProperty("finalMember", Hepsi).GetValue(dataprocessing, null) as IDictionary;
                if (am == null || fm == null || fm.Count == 0) return;

                Type ft = null;
                PropertyInfo[] k = null;
                string[] adlar = { "start_x", "start_y", "start_z", "end_x", "end_y", "end_z" };
                var geo = new HashSet<string>();
                foreach (DictionaryEntry de in fm)
                {
                    if (de.Value == null) continue;
                    if (ft == null) { ft = de.Value.GetType(); k = adlar.Select(a => ft.GetProperty(a, Hepsi)).ToArray(); }
                    geo.Add(Anahtar(Koordinat(de.Value, k)));
                }
                if (ft == null) return;

                var eklenecek = new List<KeyValuePair<string, object>>();
                double[,] aynalar = { { 1, -1 } };   // yalnız X (y -> -y)
                string[] ekler = { "X" };
                foreach (DictionaryEntry de in am)
                {
                    object a = de.Value;
                    if (a == null) continue;
                    PropertyInfo sp = a.GetType().GetProperty("SymmetricCode", Hepsi);
                    int kod = sp != null ? Convert.ToInt32(sp.GetValue(a, null)) : 0;
                    if (Array.IndexOf(DuzeltilenKodlar, kod) < 0) continue;
                    string ad = Convert.ToString(de.Key, CultureInfo.InvariantCulture);
                    if (!fm.Contains(ad)) continue;
                    object asil = fm[ad];
                    double[] p = Koordinat(asil, k);
                    for (int i = 0; i < ekler.Length; i++)
                    {
                        double fx = aynalar[i, 0], fy = aynalar[i, 1];
                        double[] q = { fx * p[0], fy * p[1], p[2], fx * p[3], fy * p[4], p[5] };
                        string yeni = ad + ekler[i];
                        if (fm.Contains(yeni) || !geo.Add(Anahtar(q))) continue;
                        object kopya = Activator.CreateInstance(ft);
                        foreach (PropertyInfo pi in ft.GetProperties(Hepsi))
                            if (pi.CanRead && pi.CanWrite && pi.GetIndexParameters().Length == 0)
                                pi.SetValue(kopya, pi.GetValue(asil, null), null);
                        for (int j = 0; j < 6; j++) k[j].SetValue(kopya, q[j], null);
                        eklenecek.Add(new KeyValuePair<string, object>(yeni, kopya));
                    }
                }
                foreach (var kv in eklenecek) fm[kv.Key] = kv.Value;
                Eklenen = eklenecek.Count;
            }
            catch { /* düzeltme yapılamazsa OutlineDrawing eskisi gibi devam eder */ }
        }

        private static double[] Koordinat(object v, PropertyInfo[] k)
        {
            var r = new double[6];
            for (int i = 0; i < 6; i++) r[i] = Convert.ToDouble(k[i].GetValue(v, null));
            return r;
        }

        private static string Anahtar(double[] q)
        {
            Func<double, string> R = v => (Math.Round(v) + 0.0).ToString(CultureInfo.InvariantCulture);
            string a = R(q[0]) + "," + R(q[1]) + "," + R(q[2]), b = R(q[3]) + "," + R(q[4]) + "," + R(q[5]);
            return string.CompareOrdinal(a, b) < 0 ? a + "|" + b : b + "|" + a;
        }
    }

    /// <summary>
    /// OutlineDrawing.SectionDetection bütün kesitlerin elemanlarını tek sözlükte eleman adıyla tutuyordu
    /// (section[elemanAdı]); bir eleman iki kesite girerse yalnız son kesitte çiziliyordu (örn. F kesitinin
    /// iki H5 elemanı G'ye gidiyordu). Tek DLL derlemesinde SectionDetection'daki üç anahtar yeri bu metotla
    /// "kesit§eleman" yapılır (eklenti/yama/OutlineDrawingYama.cs). Section.name (etiket) eleman adı kalır.
    /// </summary>
    public static class KesitAnahtari
    {
        public const char Ayrac = '\u00A7';   // §

        public static string Olustur(string uyeAnahtari, string kesitAdi)
        {
            return (kesitAdi ?? "") + Ayrac + uyeAnahtari;
        }
    }

    /// <summary>.tow dosyasında OutlineDrawing'in desteklemediği durumları bulur.</summary>
    public static class TowKontrol
    {
        /// <summary>Eleman simetri kodu 0-3 dışında olan elemanlar. OutlineDrawing (FinalizeMember) yalnız
        /// 1 (X), 2 (Y), 3 (XY) kodlarını çoğaltır; diğerlerinin aynaları çizilmez.</summary>
        public static Dictionary<int, List<string>> DesteklenmeyenSimetri(string towPath)
        {
            var sonuc = new Dictionary<int, List<string>>();
            string[] lines = File.ReadAllLines(towPath, System.Text.Encoding.GetEncoding(28591));
            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].IndexOf("; Angle Member Connectivity", StringComparison.Ordinal) < 0) continue;
                int n;
                if (!int.TryParse(lines[i].Trim().Split(' ')[0], out n)) break;
                int k = i + 1;
                for (int m = 0; m < n && k + 5 < lines.Length; m++, k += 8)
                {
                    int kod;
                    string[] t = lines[k + 5].Trim().Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                    if (t.Length == 0 || !int.TryParse(t[0], out kod) || kod <= 3) continue;
                    List<string> l;
                    if (!sonuc.TryGetValue(kod, out l)) { l = new List<string>(); sonuc[kod] = l; }
                    string grup = lines[k + 4].Trim();
                    int q = grup.IndexOf('\'');
                    int q2 = q >= 0 ? grup.IndexOf('\'', q + 1) : -1;
                    l.Add(lines[k].Trim() + (q2 > q ? " (" + grup.Substring(q + 1, q2 - q - 1) + ")" : ""));
                }
                break;
            }
            return sonuc;
        }

        /// <returns>Kullanıcıya gösterilecek uyarı, sorun yoksa null</returns>
        public static string Uyari(string towPath)
        {
            try
            {
                if (string.IsNullOrEmpty(towPath) || !File.Exists(towPath)) return null;
                var d = DesteklenmeyenSimetri(towPath);
#if TEK_DLL
                // Tek DLL'de FinalizeMember düzeltildi (SimetriDuzeltici): bu kodların aynaları artık çiziliyor
                foreach (int kod in SimetriDuzeltici.DuzeltilenKodlar)
                {
                    List<string> l;
                    if (!d.TryGetValue(kod, out l)) continue;
                    Komutlar.Yaz(string.Format("\nBilgi: simetri kodu {0} olan {1} elemanın aynaları eklendi (OutlineDrawing düzeltmesi).\n", kod, l.Count));
                    d.Remove(kod);
                }
#endif
                if (d.Count == 0) return null;
                int toplam = d.Values.Sum(l => l.Count);
                string kodlar = string.Join(", ", d.Keys.OrderBy(x => x));
                string ornek = string.Join(", ", d.Values.SelectMany(l => l).Take(4));
                Komutlar.Yaz(string.Format("\nUYARI: .tow'da simetri kodu {0} olan {1} eleman var; OutlineDrawing bunların aynalarını çizmiyor " +
                    "(çizimde eksik eleman olur). Elemanlar: {2}\n", kodlar, toplam,
                    string.Join(", ", d.Values.SelectMany(l => l))));
                return string.Format("DİKKAT: {0} elemanın simetri kodu {1}; aynaları çizilmiyor ({2}...)", toplam, kodlar, ornek);
            }
            catch { return null; }
        }
    }

    /// <summary>OutlineDrawing'in .tow okumasından gelen elemanlar ve ön/yan görünüşte görünen eleman anahtarları.</summary>
    public class KuleVerisi
    {
        public Dictionary<string, FinalMember> Uyeler;
        public List<string> GorunenAnahtarlar;
        /// <summary>Açıklaması "Redundant" içeren gruplar (grpLabel.description).</summary>
        public HashSet<string> RedundantGruplar = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        /// <summary>Ön görünüş (frontFace) ve yan görünüş (sideFace) çizgileri, OutlineDrawing'in çizdiği koordinatlarla.</summary>
        public List<YuzCizgisi> OnYuz = new List<YuzCizgisi>(), YanYuz = new List<YuzCizgisi>();
    }

    public class YuzCizgisi
    {
        public double[] P;      // { sx, sy, sz, ex, ey, ez }
        public string Grup;
        public string Kesim;    // PLS-TOWER section (Basic Body 1, 35.0 Leg Ext. ...)
    }

    /// <summary>Ön/yan görünüş çizgilerini koordinatından tanır. OutlineDrawing görünüşü
    /// Point3d(x + offset, y, z) ile çizer (ön: offset 0, yan: ön de çizildiyse 50000);
    /// "2D Aktar" sonrası hali X ekseni etrafında -90° döndürülmüştür: (x, z, -y).
    /// Arama, orta noktanın 10 mm'lik hücresi ve komşularında, uçlar 0.5 mm toleransla yapılır.</summary>
    public class YuzEslestirici
    {
        public static readonly int[] Ofsetler = { 0, 50000 };
        private const double Hucre = 10, Tol = 0.5;
        private readonly Dictionary<string, List<KeyValuePair<double[], string>>> _idx = new Dictionary<string, List<KeyValuePair<double[], string>>>();

        public YuzEslestirici(IEnumerable<YuzCizgisi> cizgiler) : this(cizgiler, Ofsetler, true, true) { }

        /// <param name="ucBoyutlu">OutlineDrawing'in çizdiği hal</param>
        /// <param name="ikiBoyutlu">2D Aktar sonrası hal</param>
        public YuzEslestirici(IEnumerable<YuzCizgisi> cizgiler, int[] ofsetler, bool ucBoyutlu, bool ikiBoyutlu)
        {
            foreach (var c in cizgiler)
                foreach (int ofs in ofsetler)
                {
                    double o = ofs;
                    double ax = c.P[0] + o, bx = c.P[3] + o;   // OutlineDrawing ile aynı toplama
                    if (ucBoyutlu) Ekle(new[] { ax, c.P[1], c.P[2], bx, c.P[4], c.P[5] }, c.Grup);
                    if (ikiBoyutlu) Ekle(new[] { ax, c.P[2], -c.P[1], bx, c.P[5], -c.P[4] }, c.Grup);   // 2D Aktar sonrası
                }
        }

        private static long G(double v) { return (long)Math.Floor(v / Hucre); }
        private static string K(long x, long y, long z) { return x + "," + y + "," + z; }

        private void Ekle(double[] p, string grup)
        {
            string k = K(G((p[0] + p[3]) / 2), G((p[1] + p[4]) / 2), G((p[2] + p[5]) / 2));
            List<KeyValuePair<double[], string>> l;
            if (!_idx.TryGetValue(k, out l)) { l = new List<KeyValuePair<double[], string>>(); _idx[k] = l; }
            l.Add(new KeyValuePair<double[], string>(p, grup));
        }

        public bool Bul(double ax, double ay, double az, double bx, double by, double bz, out string grup)
        {
            long gx = G((ax + bx) / 2), gy = G((ay + by) / 2), gz = G((az + bz) / 2);
            for (long i = gx - 1; i <= gx + 1; i++)
                for (long j = gy - 1; j <= gy + 1; j++)
                    for (long k = gz - 1; k <= gz + 1; k++)
                    {
                        List<KeyValuePair<double[], string>> l;
                        if (!_idx.TryGetValue(K(i, j, k), out l)) continue;
                        foreach (var kv in l)
                        {
                            double[] p = kv.Key;
                            bool ayni = (Y(p[0], ax) && Y(p[1], ay) && Y(p[2], az) && Y(p[3], bx) && Y(p[4], by) && Y(p[5], bz)) ||
                                        (Y(p[0], bx) && Y(p[1], by) && Y(p[2], bz) && Y(p[3], ax) && Y(p[4], ay) && Y(p[5], az));
                            if (ayni) { grup = kv.Value; return true; }
                        }
                    }
            grup = null;
            return false;
        }

        private static bool Y(double a, double b) { return Math.Abs(a - b) <= Tol; }
    }

    /// <summary>OutlineDrawing'in kendi .tow okuyucusunu (dataprocessing.Run) çağırıp sonuçları kopyalar.</summary>
    public static class KuleOkuyucu
    {
        private const BindingFlags Hepsi = BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;

        private static string _onbellekAnahtar;
        private static KuleVerisi _onbellek;

        /// <summary>Aynı .tow (yol + değişme zamanı + boyut) için son okuma tekrar kullanılır.</summary>
        public static KuleVerisi Oku(Assembly outlineDrawing, string towPath)
        {
            string anahtar = null;
            try
            {
                var fi = new FileInfo(towPath);
                anahtar = fi.FullName + "|" + fi.LastWriteTimeUtc.Ticks + "|" + fi.Length;
            }
            catch { }
            if (anahtar != null && anahtar == _onbellekAnahtar && _onbellek != null) return _onbellek;
            KuleVerisi k = OkuDogrudan(outlineDrawing, towPath);
            if (k != null && anahtar != null) { _onbellekAnahtar = anahtar; _onbellek = k; }
            return k;
        }

        public static KuleVerisi OkuDogrudan(Assembly outlineDrawing, string towPath)
        {
            Type dp = outlineDrawing.GetType("OutlineDrawing.dataprocessing", true);
            MethodInfo run = dp.GetMethod("Run", Hepsi, null, new[] { typeof(string) }, null);
            if (run == null) throw new MissingMethodException("OutlineDrawing.dataprocessing", "Run");
            object data = run.Invoke(null, new object[] { towPath });
            if (data == null) return null;
            return new KuleVerisi
            {
                Uyeler = Kopyala(data),
                GorunenAnahtarlar = Gorunenler(data),
                RedundantGruplar = Redundantlar(data, towPath),
                OnYuz = Yuz(data, "frontFace"),
                YanYuz = Yuz(data, "sideFace")
            };
        }

        /// <summary>frontFace / sideFace sözlüğündeki çizgiler (OutlineDrawing'in çizdiği değerler).</summary>
        public static List<YuzCizgisi> Yuz(object data, string ad)
        {
            var sonuc = new List<YuzCizgisi>();
            PropertyInfo pi = data.GetType().GetProperty(ad, Hepsi);
            var d = pi != null ? pi.GetValue(data, null) as IDictionary : null;
            if (d == null) return sonuc;
            string[] adlar = { "start_x", "start_y", "start_z", "end_x", "end_y", "end_z" };
            PropertyInfo[] p = null; PropertyInfo g = null, ks = null;
            foreach (DictionaryEntry de in d)
            {
                object v = de.Value;
                if (v == null) continue;
                if (p == null)
                {
                    p = adlar.Select(a => v.GetType().GetProperty(a, Hepsi)).ToArray();
                    g = v.GetType().GetProperty("group_label", Hepsi);
                    ks = v.GetType().GetProperty("section_label", Hepsi);
                }
                sonuc.Add(new YuzCizgisi
                {
                    P = p.Select(x => Convert.ToDouble(x.GetValue(v, null))).ToArray(),
                    Grup = g != null ? g.GetValue(v, null) as string : null,
                    Kesim = ks != null ? ks.GetValue(v, null) as string : null
                });
            }
            return sonuc;
        }

        /// <summary>Redundant gruplar: .tow grup tablosunda PLS-TOWER grup tipi Redundant olanlar
        /// ve açıklamasında "redundant" geçenler.</summary>
        public static HashSet<string> Redundantlar(object data, string towPath)
        {
            HashSet<string> sonuc = Redundantlar(data);
            try { sonuc.UnionWith(TowRedundantGruplari(towPath)); }
            catch (System.Exception ex) { System.Diagnostics.Debug.WriteLine("Grup tipi okunamadı: " + ex.Message); }
            return sonuc;
        }

        /// <summary>PLS-TOWER grup tipi kodu: 1 Leg, 2 Other, 3 Redundant (PLS-TOWER çiziminde turuncu).</summary>
        public const int RedundantGrupTipi = 3;

        /// <summary>.tow grup tablosunda ("; group label, description, size, material, angle type, element type, group type")
        /// grup tipi Redundant olan grup adları. OutlineDrawing bu sütunu okumadığı için dosyadan okunur.</summary>
        public static HashSet<string> TowRedundantGruplari(string towPath)
        {
            var sonuc = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrEmpty(towPath) || !File.Exists(towPath)) return sonuc;
            string[] satirlar = File.ReadAllLines(towPath, System.Text.Encoding.GetEncoding(28591));
            for (int i = 0; i < satirlar.Length; i++)
            {
                string s = satirlar[i];
                int yorum = s.IndexOf(';');
                if (yorum < 0 || s.IndexOf("group label", yorum, StringComparison.OrdinalIgnoreCase) < 0) continue;
                string[] sutunlar = s.Substring(yorum + 1).Split(',').Select(x => x.Trim()).ToArray();
                int tipSutunu = Array.FindIndex(sutunlar, x => x.Equals("group type", StringComparison.OrdinalIgnoreCase));
                int adet;
                if (tipSutunu < 0 || !int.TryParse(s.Substring(0, yorum).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out adet)) continue;
                for (int k = i + 1; k <= i + adet && k < satirlar.Length; k++)
                {
                    List<string> t = Parcala(satirlar[k]);
                    int tip;
                    if (t.Count > tipSutunu && int.TryParse(t[tipSutunu], NumberStyles.Integer, CultureInfo.InvariantCulture, out tip)
                        && tip == RedundantGrupTipi)
                        sonuc.Add(t[0].Trim());
                }
                i += adet;
            }
            return sonuc;
        }

        /// <summary>Tırnaklı alanları tek parça sayan boşluk ayırıcı ('a b' 'c' 1 2 -> a b | c | 1 | 2).</summary>
        private static List<string> Parcala(string s)
        {
            var t = new List<string>();
            foreach (Match m in Regex.Matches(s, @"'([^']*)'|(\S+)"))
                t.Add(m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value);
            return t;
        }

        /// <summary>grpLabel'de açıklaması "redundant" içeren grup adları.</summary>
        public static HashSet<string> Redundantlar(object data)
        {
            var sonuc = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            PropertyInfo pi = data.GetType().GetProperty("grpLabel", Hepsi);
            var d = pi != null ? pi.GetValue(data, null) as IDictionary : null;
            if (d == null) return sonuc;
            foreach (DictionaryEntry de in d)
            {
                if (de.Value == null) continue;
                PropertyInfo dp = de.Value.GetType().GetProperty("description", Hepsi);
                string aciklama = dp != null ? dp.GetValue(de.Value, null) as string : null;
                if (aciklama != null && aciklama.IndexOf("redundant", StringComparison.OrdinalIgnoreCase) >= 0)
                    sonuc.Add(Convert.ToString(de.Key, CultureInfo.InvariantCulture).Trim().Trim('\''));
            }
            return sonuc;
        }

        /// <summary>frontFace + sideFace anahtarları (yoksa null).</summary>
        public static List<string> Gorunenler(object data)
        {
            var sonuc = new List<string>();
            bool bulundu = false;
            foreach (string ad in new[] { "frontFace", "sideFace" })
            {
                PropertyInfo pi = data.GetType().GetProperty(ad, Hepsi);
                var d = pi != null ? pi.GetValue(data, null) as IDictionary : null;
                if (d == null) continue;
                bulundu = true;
                foreach (object k in d.Keys) sonuc.Add(Convert.ToString(k, CultureInfo.InvariantCulture));
            }
            return bulundu ? sonuc : null;
        }

        /// <summary>dataprocessing nesnesindeki finalMember sözlüğünü yerel FinalMember'lara kopyalar.</summary>
        public static Dictionary<string, FinalMember> Kopyala(object data)
        {
            var fm = data.GetType().GetProperty("finalMember", Hepsi).GetValue(data, null) as IDictionary;
            if (fm == null) return null;

            var sonuc = new Dictionary<string, FinalMember>();
            PropertyInfo[] p = null;
            string[] adlar = { "start_x", "start_y", "start_z", "end_x", "end_y", "end_z", "group_label", "size", "section_label" };
            foreach (DictionaryEntry de in fm)
            {
                object v = de.Value;
                if (v == null) continue;
                if (p == null) p = adlar.Select(a => v.GetType().GetProperty(a, Hepsi)).ToArray();
                sonuc[Convert.ToString(de.Key, CultureInfo.InvariantCulture)] = new FinalMember
                {
                    start_x = Convert.ToDouble(p[0].GetValue(v, null)),
                    start_y = Convert.ToDouble(p[1].GetValue(v, null)),
                    start_z = Convert.ToDouble(p[2].GetValue(v, null)),
                    end_x = Convert.ToDouble(p[3].GetValue(v, null)),
                    end_y = Convert.ToDouble(p[4].GetValue(v, null)),
                    end_z = Convert.ToDouble(p[5].GetValue(v, null)),
                    group_label = p[6] != null ? p[6].GetValue(v, null) as string : null,
                    size = p[7] != null ? p[7].GetValue(v, null) as string : null,
                    section_label = p[8] != null ? p[8].GetValue(v, null) as string : null
                };
            }
            return sonuc;
        }
    }

    // AutoSectionDetector'un kullandığı yerel tipler (OutlineDrawing'deki aynı adlı tiplerin sade kopyası)
    public class FinalMember
    {
        public double start_x { get; set; }
        public double start_y { get; set; }
        public double start_z { get; set; }
        public double end_x { get; set; }
        public double end_y { get; set; }
        public double end_z { get; set; }
        public string group_label { get; set; }
        public string size { get; set; }
        public string material { get; set; }
        public string section_label { get; set; }
    }

    public class SectionPoint3D
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        public SectionPoint3D() { }
        public SectionPoint3D(double x, double y, double z) { X = x; Y = y; Z = z; }
    }
}
