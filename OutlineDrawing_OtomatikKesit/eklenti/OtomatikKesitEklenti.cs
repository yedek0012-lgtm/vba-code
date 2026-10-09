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
//      ana elemanlar düz/cyan, redundant elemanlar (grup açıklaması "Redundant") kesikli/mavi olur.
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
[assembly: CommandClass(typeof(OtomatikKesit.Komutlar))]

namespace OtomatikKesit
{
    public class Eklenti : IExtensionApplication
    {
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
            Komutlar.Yaz("\nOtomatik Kesit eklentisi yüklendi. OutlineDrawing formunda 'Oto Kesit' butonu görünecek (komutlar: OTOKESIT, OTOOLCU).\n");
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
            KesitOlculeri.Ekle();
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
            Yaz("\n" + KesitStili.Uygula(f) + "\n");
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
            tip.SetToolTip(chkR, "ÇALIŞTIR'dan sonra kesitlerde ana elemanlar düz, redundant elemanlar kesikli çizilsin (komut: OTOSTIL)");
            b.Click += (s, e) => TabloyuDoldur(f);

            // ÇALIŞTIR'a ikinci bir olay bağlanır; OutlineDrawing'in kendi çizimi bittikten sonra çalışır.
            var run = Alan(f, "btnRun") as Button;
            if (run != null)
                run.Click += (s, e) =>
                {
                    var mesaj = new List<string>();
                    if (chkR.Checked)
                    {
                        try { mesaj.Add(KesitStili.Uygula(f)); }
                        catch (System.Exception ex) { mesaj.Add("Redundant çizgi tipi uygulanamadı: " + ex.Message); }
                    }
                    if (chk.Checked)
                    {
                        try { mesaj.Add(KesitOlculeri.Ekle()); }
                        catch (System.Exception ex) { mesaj.Add("Ölçü eklenemedi: " + ex.Message); }
                    }
                    if (mesaj.Count > 0) Uyari(f, string.Join(" | ", mesaj));
                };
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

                string ozet = string.Format("{0} otomatik kesit eklendi.", bulunan.Count);
                if (rapor != null)
                    ozet += string.Format(" Görünmeyen {0} elemanın {1} tanesi kesitlerde.", rapor.HiddenMembers, rapor.CoveredHidden);
                Uyari(f, ozet + " Kontrol edip ÇALIŞTIR'a basın.");

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

        internal static void Uyari(Form f, string mesaj)
        {
            // OutlineDrawing'in kendi uyarı etiketi; yoksa komut satırı
            if (!Cagir(f, "PrintErrorLabel", 5000, mesaj)) Komutlar.Yaz("\n" + mesaj + "\n");
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
                    ent.UpgradeOpen();
                    ent.Erase();
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
                    if (Math.Abs(wa - e.W) > 1 || Math.Abs(wb - e.W) > 1) continue;
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
            int adet = 0, kesit = 0;

            using (doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var ms = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);
                KatmaniHazirla(db, tr);

                foreach (KesitCizimi e in KesitCizimleri.Topla(tr, ms, Katman))
                {
                    if (e.Segs.Count == 0) continue;
                    List<OlcuTanimi> olculer = OlcuPlani.Hesapla(e.Segs);
                    Matrix3d m = e.XZ
                        ? Matrix3d.Displacement(new Vector3d(0, e.W, 0)) * Matrix3d.Rotation(Math.PI / 2, Vector3d.XAxis, Point3d.Origin)
                        : Matrix3d.Displacement(new Vector3d(0, 0, e.W));
                    foreach (OlcuTanimi o in olculer)
                    {
                        var dim = new RotatedDimension(o.Rotation,
                            new Point3d(o.U1, o.V1, 0), new Point3d(o.U2, o.V2, 0), new Point3d(o.UD, o.VD, 0),
                            "", db.Dimstyle);
                        dim.SetDatabaseDefaults(db);
                        dim.Layer = Katman;
                        dim.Dimtxt = 200;
                        dim.Dimasz = 150;
                        dim.Dimexo = 100;
                        dim.Dimexe = 100;
                        dim.Dimgap = 50;
                        dim.Dimdec = 0;
                        dim.Dimtad = 1;
                        dim.Dimtih = false;
                        dim.Dimtoh = false;
                        dim.TransformBy(m);
                        ms.AppendEntity(dim);
                        tr.AddNewlyCreatedDBObject(dim, true);
                        adet++;
                    }
                    kesit++;
                }
                tr.Commit();
            }
            string ozet = string.Format("Ölçü: {0} kesite {1} ölçü (katman {2})", kesit, adet, Katman);
            Komutlar.Yaz("\n" + ozet + ".\n");
            return ozet;
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
            if (kule.RedundantGruplar.Count == 0) return "Redundant stili: .tow'da açıklaması 'Redundant' olan grup yok.";

            Dictionary<string, KesitTanimi> tanimlar = GridTanimlari(f);
            int ana = 0, red = 0, eslesmeyen = 0, kesit = 0, hata = 0, yuzAna = 0, yuzRed = 0;
            var tanimsiz = new List<string>();
            string ilkHata = null;

            using (doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var ms = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);
                ObjectId kesik = KesikCizgiTipi(db, tr);
                ObjectId duz = SymbolUtilityServices.GetLinetypeContinuousId(db);
                double olcek = db.Ltscale > 1e-9 ? 1.0 / db.Ltscale : 1.0;

                // ---- ön / yan görünüş: çizgiler koordinatından tanınır
                var yuz = new YuzEslestirici(kule.OnYuz.Concat(kule.YanYuz));
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
                        if (eslesme[i] < 0) { eslesmeyen++; continue; }
                        try
                        {
                            FinalMember m = kule.Uyeler[anahtarlar[eslesme[i]]];
                            bool redundant = m.group_label != null && kule.RedundantGruplar.Contains(m.group_label.Trim());
                            var ent = (Entity)tr.GetObject(e.Ids[i], OpenMode.ForWrite);
                            if (Boya(ent, redundant, kesik, duz, olcek)) red++; else ana++;
                        }
                        catch (System.Exception ex)
                        {
                            hata++;
                            if (ilkHata == null) ilkHata = ex.Message;
                        }
                    }
                }
                tr.Commit();
            }

            string ozet = string.Format("Çizgi tipi: görünüşlerde {0} redundant / {1} ana, {2} kesitte {3} redundant / {4} ana",
                yuzRed, yuzAna, kesit, red, ana);
            if (eslesmeyen > 0) ozet += ", eşleşmeyen çizgi " + eslesmeyen;
            if (tanimsiz.Count > 0) ozet += ", tabloda bulunamayan kesit: " + string.Join(",", tanimsiz);
            if (hata > 0) ozet += string.Format(", {0} çizgide hata ({1})", hata, ilkHata);
            Komutlar.Yaz("\n" + ozet + ".\n");
            return ozet;
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

            int enIyi = -1; int[] enIyiEslesme = null;
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
                            if (skor > enIyi) { enIyi = skor; enIyiEslesme = es; }
                            if (enIyi == hedef.Count) return enIyiEslesme;
                        }
                }
            }
            return enIyiEslesme ?? sonuc;
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
            var pts = new List<double[]>();
            foreach (var s in segs) { pts.Add(new[] { s[0], s[1] }); pts.Add(new[] { s[2], s[3] }); }
            var hull = Sadelestir(Kontur(pts), KirikAcisi);
            if (hull.Count < 2) return sonuc;

            double minU = hull.Min(p => p[0]), maxU = hull.Max(p => p[0]);
            double minV = hull.Min(p => p[1]), maxV = hull.Max(p => p[1]);

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
                double vd = maxV + Aralik;
                for (int i = 0; i + 1 < us.Count; i++)
                    sonuc.Add(new OlcuTanimi
                    {
                        U1 = us[i][0], V1 = us[i][1], U2 = us[i + 1][0], V2 = us[i + 1][1],
                        UD = (us[i][0] + us[i + 1][0]) / 2, VD = vd, Rotation = 0
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
                    UD = maxU + Aralik, VD = (maxV + minV) / 2, Rotation = Math.PI / 2
                });
            }
            return sonuc;
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

        public YuzEslestirici(IEnumerable<YuzCizgisi> cizgiler)
        {
            foreach (var c in cizgiler)
                foreach (int ofs in Ofsetler)
                {
                    double o = ofs;
                    double ax = c.P[0] + o, bx = c.P[3] + o;   // OutlineDrawing ile aynı toplama
                    Ekle(new[] { ax, c.P[1], c.P[2], bx, c.P[4], c.P[5] }, c.Grup);
                    Ekle(new[] { ax, c.P[2], -c.P[1], bx, c.P[5], -c.P[4] }, c.Grup);   // 2D Aktar sonrası
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

        public static KuleVerisi Oku(Assembly outlineDrawing, string towPath)
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
                RedundantGruplar = Redundantlar(data),
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
            PropertyInfo[] p = null; PropertyInfo g = null;
            foreach (DictionaryEntry de in d)
            {
                object v = de.Value;
                if (v == null) continue;
                if (p == null) { p = adlar.Select(a => v.GetType().GetProperty(a, Hepsi)).ToArray(); g = v.GetType().GetProperty("group_label", Hepsi); }
                sonuc.Add(new YuzCizgisi
                {
                    P = p.Select(x => Convert.ToDouble(x.GetValue(v, null))).ToArray(),
                    Grup = g != null ? g.GetValue(v, null) as string : null
                });
            }
            return sonuc;
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
