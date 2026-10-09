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
//
//  OutlineDrawing'in iç tiplerine yansıtma (reflection) ile erişilir, bu yüzden
//  OutlineDrawing.dll'in yeniden derlenmesi gerekmez.
//  Komutlar: OTOKESIT (tabloyu doldur), OTOOLCU (çizili kesitlere ölçü koy / yenile)
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
            }
            else
            {
                b.Size = new Size(70, 24);
                b.Location = new Point(8, 8);
                chk.Location = new Point(8, 36);
            }

            kap.Controls.Add(b);
            kap.Controls.Add(chk);
            b.BringToFront();
            chk.BringToFront();
            var tip = new ToolTip();
            tip.SetToolTip(b, "Ön/yan görünüşte görünmeyen elemanların hepsini kapsayan kesitleri bulup tabloya yazar");
            tip.SetToolTip(chk, "ÇALIŞTIR'dan sonra kesitlere ölçü koy (komut: OTOOLCU)");
            b.Click += (s, e) => TabloyuDoldur(f);

            // ÇALIŞTIR'a ikinci bir olay bağlanır; OutlineDrawing'in kendi çizimi bittikten sonra çalışır.
            var run = Alan(f, "btnRun") as Button;
            if (run != null)
                run.Click += (s, e) =>
                {
                    if (!chk.Checked) return;
                    try { KesitOlculeri.Ekle(); }
                    catch (System.Exception ex) { Komutlar.Yaz("\nÖlçü eklenemedi: " + ex.Message + "\n"); }
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

        private static void Uyari(Form f, string mesaj)
        {
            // OutlineDrawing'in kendi uyarı etiketi; yoksa komut satırı
            if (!Cagir(f, "PrintErrorLabel", 5000, mesaj)) Komutlar.Yaz("\n" + mesaj + "\n");
        }

        private static object Alan(object o, string ad)
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
    //  Kesit ölçüleri
    // =================================================================
    internal static class KesitOlculeri
    {
        internal const string Katman = "OTO_KESIT_OLCU";
        private static readonly Regex KesitYazisi = new Regex(@"^SECTION\s+\S+$", RegexOptions.IgnoreCase);

        private class Etiket { public double U, V, W; public bool XZ; public string Ad; }

        /// <summary>Model alanındaki her "SECTION x" çizimine ölçü koyar; önceki otomatik ölçüleri siler.</summary>
        internal static void Ekle()
        {
            Document doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            int adet = 0, kesit = 0;

            using (doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var ms = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);
                KatmaniHazirla(db, tr);

                var etiketler = new List<Etiket>();
                var cizgiler = new List<Point3d[]>();
                foreach (ObjectId id in ms)
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent == null) continue;
                    if (string.Equals(ent.Layer, Katman, StringComparison.OrdinalIgnoreCase))
                    {
                        ent.UpgradeOpen();
                        ent.Erase();
                        continue;
                    }
                    var t = ent as DBText;
                    if (t != null)
                    {
                        if (t.TextString == null || !KesitYazisi.IsMatch(t.TextString.Trim()) || t.Height < 250) continue;
                        Vector3d n = t.Normal;
                        Point3d p = t.HorizontalMode == TextHorizontalMode.TextLeft && t.VerticalMode == TextVerticalMode.TextBase
                            ? t.Position : t.AlignmentPoint;
                        if (Math.Abs(n.Y) > 0.9) etiketler.Add(new Etiket { U = p.X, V = p.Z, W = p.Y, XZ = true, Ad = t.TextString.Trim() });
                        else if (Math.Abs(n.Z) > 0.9) etiketler.Add(new Etiket { U = p.X, V = p.Y, W = p.Z, XZ = false, Ad = t.TextString.Trim() });
                        continue;
                    }
                    var l = ent as Line;
                    if (l != null && !l.Layer.StartsWith("SECTION_MARK", StringComparison.OrdinalIgnoreCase))
                        cizgiler.Add(new[] { l.StartPoint, l.EndPoint });
                }

                foreach (Etiket e in etiketler)
                {
                    // Bu etikete ait çizgiler: aynı düzlemde, en yakın etiket bu olan ve 40 m içinde kalanlar
                    var segs = new List<double[]>();
                    foreach (Point3d[] c in cizgiler)
                    {
                        double ua = c[0].X, ub = c[1].X;
                        double va = e.XZ ? c[0].Z : c[0].Y, vb = e.XZ ? c[1].Z : c[1].Y;
                        double wa = e.XZ ? c[0].Y : c[0].Z, wb = e.XZ ? c[1].Y : c[1].Z;
                        if (Math.Abs(wa - e.W) > 1 || Math.Abs(wb - e.W) > 1) continue;
                        double um = (ua + ub) / 2, vm = (va + vb) / 2;
                        if (Math.Abs(um - e.U) > 15000 || Math.Abs(vm - e.V) > 40000) continue;
                        bool baskasi = etiketler.Any(o => o != e && o.XZ == e.XZ && Math.Abs(o.W - e.W) <= 1 &&
                                                          Math.Abs(um - o.U) < Math.Abs(um - e.U) && Math.Abs(vm - o.V) <= 40000);
                        if (baskasi) continue;
                        segs.Add(new[] { ua, va, ub, vb });
                    }
                    if (segs.Count == 0) continue;

                    List<OlcuTanimi> olculer = OlcuPlani.Hesapla(segs);
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
            Komutlar.Yaz(string.Format("\nKesit ölçüleri: {0} kesite {1} ölçü eklendi (katman {2}).\n", kesit, adet, Katman));
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
            return new KuleVerisi { Uyeler = Kopyala(data), GorunenAnahtarlar = Gorunenler(data) };
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
