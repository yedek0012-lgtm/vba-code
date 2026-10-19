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
        public const string Surum = "3.0";
        private static Timer _timer;

        public void Initialize()
        {
            // Modeless OutlineDrawing formunu yakalamak için UI iş parçacığında küçük bir zamanlayıcı
            _timer = new Timer { Interval = 700 };
            _timer.Tick += (s, e) =>
            {
                try { FormBaglayici.FormlariTara(); }
                catch { }
                try { KesikIzleyici.Tik(); }
                catch { }
            };
            _timer.Start();
            Komutlar.Yaz("\nOtomatik Kesit " + Surum + " yüklendi. OutlineDrawing formunda 'Oto Kesit' butonu görünecek (komutlar: OTOKESIT, OTOOLCU, OTOSTIL, OTOIC).\n");
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
            FormBaglayici.Guvenli(null, "Ölçü", () => { KesitOlculeri.Ekle(); Goruntu.Yenile(); });
        }

        [CommandMethod("OTOIC", CommandFlags.Session)]
        public void OtoIc()
        {
            Form f = FormBaglayici.AcikFormuBul();
            if (f == null)
            {
                Yaz("\nOutlineDrawing formu açık değil. Önce OutlineDrawing'i çalıştırıp .tow dosyasını seçin.\n");
                return;
            }
            FormBaglayici.Guvenli(f, "İç eleman", () => { FormBaglayici.Uyari(f, IcElemanlar.Ciz(f, true)); Goruntu.Yenile(); });
        }

        [CommandMethod("OTOISARET", CommandFlags.Session)]
        public void OtoIsaret()
        {
            Form f = FormBaglayici.AcikFormuBul();
            if (f == null)
            {
                Yaz("\nOutlineDrawing formu açık değil. Önce OutlineDrawing'i çalıştırıp .tow dosyasını seçin.\n");
                return;
            }
            FormBaglayici.Guvenli(f, "Kesit işaretleri", () => { Yaz("\n" + (KesitIsaretleri.Ciz(f) ?? "Kesit işareti yok (tablo boş ya da ön görünüş bulunamadı).") + "\n"); Goruntu.Yenile(); });
        }

        [CommandMethod("OTOPAFTA", CommandFlags.Session)]
        public void OtoPafta()
        {
            Form f = FormBaglayici.AcikFormuBul();
            try
            {
                string sablon = Ayarlar.PaftaSablonu != null && File.Exists(Ayarlar.PaftaSablonu) ? Ayarlar.PaftaSablonu : null;
                Yaz("\n" + Pafta.Olustur(f, sablon) + "\n");
                Goruntu.Yenile();
            }
            catch (System.Exception ex) { Yaz("\nPafta kurulamadı: " + FormBaglayici.IcHata(ex).Message + "\n"); }
        }

        [CommandMethod("OTODIZ", CommandFlags.Session)]
        public void OtoDiz()
        {
            try
            {
                string r = KesitDizici.Diz();
                Yaz("\n" + (r ?? "Dizilecek kesit yok (en az iki \"SECTION x\" başlığı gerekir).") + "\n");
                Goruntu.Yenile();
            }
            catch (System.Exception ex) { Yaz("\nKesitler dizilemedi: " + ex.Message + "\n"); }
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
            FormBaglayici.Guvenli(f, "Çizgi tipi", () => { Yaz("\n" + KesitStili.Uygula(f) + "\n"); Goruntu.Yenile(); });
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

        private static readonly Dictionary<string, DateTime> _sonYazilan = new Dictionary<string, DateTime>();

        /// <summary>Aynı mesaj kısa süre içinde (bir ÇALIŞTIR boyunca birkaç adım aynı bilgiyi üretir) bir kez yazılır.</summary>
        internal static void BirKezYaz(string msg)
        {
            DateTime son, simdi = DateTime.UtcNow;
            if (_sonYazilan.TryGetValue(msg, out son) && (simdi - son).TotalSeconds < 10) return;
            _sonYazilan[msg] = simdi;
            if (_sonYazilan.Count > 200) _sonYazilan.Clear();
            Yaz(msg);
        }
    }

    internal static class FormBaglayici
    {
        private const string FormTipi = "OutlineDrawing.mainform";
        private const string ButonAdi = "btnAutoSection";
        private const string HepsiButonAdi = "btnHepsiniCiz";
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
            // Ekran ölçeği / monitör değişince form küçük ya da boş görünüyordu: her açılışta ekranı kaplasın
            try { f.WindowState = FormWindowState.Maximized; } catch { }
            var ornek = f.Controls.Find("btnHeight", true).FirstOrDefault() as Button;
            Control kap = ornek != null ? ornek.Parent : (Control)f;

            var b = new Button { Name = ButonAdi, Text = "Oto Kesit" };
            var bh = new Button { Name = HepsiButonAdi, Text = "Hepsini Çiz" };
            var bp = new Button { Name = "btnPafta", Text = "Pafta" };
            var chk = new CheckBox { Name = OlcuKutusuAdi, Text = "Ölçü", Checked = true, AutoSize = true };
            var chkR = new CheckBox { Name = "chkRedundant", Text = "Redundant", Checked = true, AutoSize = true };
            var chkI = new CheckBox { Name = "chkIcEleman", Text = "İç eleman", Checked = true, AutoSize = true };
            var chkD = new CheckBox { Name = "chkKesitDiz", Text = "Kesitleri diz", Checked = true, AutoSize = true };
            var chkT = new CheckBox { Name = TumElemanKutusuAdi, Text = "Tüm elemanlar", Checked = false, AutoSize = true };
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
                foreach (Button bb in new[] { bh, bp })
                {
                    bb.Size = ornek.Size; bb.Font = ornek.Font; bb.FlatStyle = ornek.FlatStyle;
                    bb.BackColor = ornek.BackColor; bb.ForeColor = ornek.ForeColor; bb.UseVisualStyleBackColor = ornek.UseVisualStyleBackColor;
                    bb.Cursor = ornek.Cursor; bb.Anchor = AnchorStyles.Top | AnchorStyles.Left;
                    bb.FlatAppearance.BorderColor = ornek.FlatAppearance.BorderColor; bb.FlatAppearance.BorderSize = ornek.FlatAppearance.BorderSize;
                }
                bh.Font = new System.Drawing.Font(ornek.Font, FontStyle.Bold);
                b.Anchor = AnchorStyles.Top | AnchorStyles.Left;
                foreach (CheckBox c in new[] { chk, chkR, chkI, chkD, chkT }) { c.Font = ornek.Font; c.ForeColor = kap.ForeColor; c.Anchor = AnchorStyles.Top | AnchorStyles.Left; }
                // Yer: ÇALIŞTIR'ın sağı, tablonun ve sağ buton sütununun (Nokta ... Liste) altı. Soldan sağa dizilir, kutunun
                // sağ kenarına gelince alt satıra geçer (önceden kutucuklar sağ kenarın dışına düşüp görünmüyordu).
                Control calistir = kap.Controls.Find("btnRun", false).FirstOrDefault();
                Control izgara = kap.Controls.Find("dataGridView1", false).FirstOrDefault();
                int sutunAlt = kap.Controls.Cast<Control>().Where(c => c is Button && Math.Abs(c.Left - ornek.Left) <= 3).Max(c => c.Bottom);
                int x0 = calistir != null ? calistir.Right + 10 : ornek.Left;
                int y0 = Math.Max(sutunAlt, izgara != null ? izgara.Bottom : 0) + 6;
                int sag = kap.ClientSize.Width - 6, px = x0, py = y0, satirH = 0;
                Action<Control, bool> koy = (c, yeniSatir) =>
                {
                    Size sz = c is CheckBox ? c.PreferredSize : c.Size;
                    if (yeniSatir && px > x0) { px = x0; py += satirH + 4; satirH = 0; }
                    if (px > x0 && px + sz.Width > sag) { px = x0; py += satirH + 4; satirH = 0; }
                    c.Location = new Point(px, py);
                    px += sz.Width + 6; satirH = Math.Max(satirH, sz.Height);
                };
                koy(b, false); koy(bh, false); koy(bp, false);
                koy(chk, true); koy(chkR, false); koy(chkI, false); koy(chkD, false); koy(chkT, false);
            }
            else
            {
                b.Size = new Size(70, 24);
                b.Location = new Point(8, 8);
                bh.Size = new Size(90, 24);
                bh.Location = new Point(84, 8);
                bp.Size = new Size(60, 24);
                bp.Location = new Point(180, 8);
                chk.Location = new Point(8, 36);
                chkR.Location = new Point(8, 58);
                chkI.Location = new Point(8, 80);
                chkD.Location = new Point(8, 102);
                chkT.Location = new Point(8, 124);
            }

            kap.Controls.Add(b);
            kap.Controls.Add(bh);
            kap.Controls.Add(bp);
            kap.Controls.Add(chk);
            kap.Controls.Add(chkR);
            kap.Controls.Add(chkI);
            kap.Controls.Add(chkD);
            kap.Controls.Add(chkT);
            b.BringToFront();
            bh.BringToFront();
            bp.BringToFront();
            chk.BringToFront();
            chkR.BringToFront();
            chkI.BringToFront();
            chkD.BringToFront();
            chkT.BringToFront();
            var tip = new ToolTip();
            tip.SetToolTip(b, "Çizimci gibi sade kesitleri bulup tabloya yazar: yatay planlar, travers düzlemleri, çerçeveler, bacak iç düzlemleri");
            tip.SetToolTip(chkT, "İşaretliyse her gizli eleman mutlaka bir kesitte gösterilir (küçük, eğik parça kesitler de alınır). " +
                                 "İşaretsizse sade: yalnız mantıklı yerlerden kesit, kalan elemanlar görünüşte iç eleman çizgisi olarak durur");
            tip.SetToolTip(chk, "ÇALIŞTIR'dan sonra kesitlere ölçü koy (komut: OTOOLCU)");
            tip.SetToolTip(chkR, "YÜKLE ve ÇALIŞTIR'dan sonra 3D modelde, görünüşlerde ve kesitlerde ana elemanlar düz, redundant elemanlar kesikli; isimler tek renk (komut: OTOSTIL)");
            tip.SetToolTip(chkD, "ÇALIŞTIR'dan sonra kesitleri kendi boyutlarına göre satırlar halinde sıkı bir ızgaraya dizer (tek uzun sıra yerine); çizgiler, etiketler, başlık ve ölçüler birlikte taşınır (komut: OTODIZ)");
            tip.SetToolTip(chkI, "ÇALIŞTIR'dan sonra ön/yan görünüşe yüzde olmayan iç elemanların (kalça çaprazları, diyaframlar) izdüşümünü PLS-TOWER'daki gibi ekler; katman " + IcElemanlar.Katman + " (komut: OTOIC)");
            b.Click += (s, e) => Guvenli(f, "Otomatik kesit", () => TabloyuDoldur(f));
            tip.SetToolTip(bh, "Tek tık: otomatik kesitleri bulup tabloya yazar ve ÇALIŞTIR'ı çalıştırır (görünüşler, kesitler, iç eleman, ölçü, redundant)");
            bh.Click += (s, e) => Guvenli(f, "Hepsini Çiz", () =>
            {
                Ayarlar.Kaydet(f);
                // Tablo bu .tow için Oto Kesit'in yazdığı haliyle duruyorsa yeniden hesaplanıp "silinsin mi?" diye sorulmaz
                if (!TabloOtoHalinde(f))
                {
                    TabloyuDoldur(f);
                    if (!_sonDoldurmaTamam) return;
                }
                var calistir = Alan(f, "btnRun") as Button;
                if (calistir != null) calistir.PerformClick();
            });

            tip.SetToolTip(bp, "ÇALIŞTIR'dan sonra: görünüşleri, kesitleri ve .tow'dan malzeme listesini pafta (antet) çerçevesine yerleştirir; ölçek otomatik. Ctrl+tık: pafta şablonu seç (DWG/DXF; varsayılan ŞA-RA A1) (komut: OTOPAFTA)");
            bp.Click += (s, e) => Guvenli(f, "Pafta", () =>
            {
                if ((Control.ModifierKeys & Keys.Control) != 0)
                {
                    using (var dlg = new OpenFileDialog { Filter = "Pafta şablonu (*.dwg;*.dxf)|*.dwg;*.dxf", Title = "Pafta şablonu (çerçeve + antet; sol alt köşe 0,0)" })
                        if (dlg.ShowDialog(f) == DialogResult.OK) Ayarlar.PaftaSablonu = dlg.FileName; else return;
                }
                string sablon = Ayarlar.PaftaSablonu != null && File.Exists(Ayarlar.PaftaSablonu) ? Ayarlar.PaftaSablonu : null;
                string r = Pafta.Olustur(f, sablon);
                Goruntu.Yenile();
                Koordinat.Kesit();
                Ayarlar.Kaydet(f);
                Uyari(f, r);
            });

            // Son kullanılan .tow, görünüş / metin seçenekleri ve kutucuklar hatırlanır
            Ayarlar.Yukle(f);
            f.FormClosing += (s, e) => Ayarlar.Kaydet(f);

            // Kesit satırı silinince / içi boşaltılınca harfler (A, B, C ...) yeniden sıralanır. OutlineDrawing
            // harfleri yalnız Nokta / Yükseklik ile satır eklenince yeniliyordu.
            var tablo = Alan(f, "dataGridView1") as DataGridView;
            if (tablo != null)
            {
                Action yenile = () =>
                {
                    if (_harfYenileniyor || f.IsDisposed) return;
                    try { f.BeginInvoke((Action)(() => Guvenli(f, "Kesit harfleri", () => HarfleriYenile(f)))); } catch { }
                };
                tablo.RowsRemoved += (s, e) => { if (!_tabloYaziliyor) _otoTabloAnahtari = null; yenile(); };
                tablo.CellValueChanged += (s, e) =>
                {
                    if (e.ColumnIndex < 0 || tablo.Columns[e.ColumnIndex].Name == "ColumnSection") return;
                    if (!_tabloYaziliyor) _otoTabloAnahtari = null;
                    yenile();
                };
            }

            // YÜKLE ve ÇALIŞTIR model alanındaki HER ŞEYİ siler: başka katmanda çizim varsa önce sor
            var load = Alan(f, "btnLoad") as Button;
            var run = Alan(f, "btnRun") as Button;
            var iki = Alan(f, "btnConvertto2D") as Button;
            // 2D Aktar her şeyi XY düzlemine çevirir: UCS de dünya koordinatına döner
            if (iki != null) iki.Click += (s, e) => Guvenli(f, "2D Aktar sonrası", () => Koordinat.Dunya());
            bool loadSarili = SilmeKorumasiEkle(f, load, "YÜKLE");
            bool runSarili = SilmeKorumasiEkle(f, run, "ÇALIŞTIR");
            DunyaUcsIleSec(Alan(f, "btnSelectMultiplePoint") as Button);
            DunyaUcsIleSec(Alan(f, "btnHeight") as Button);

            // YÜKLE (3D model) sonrasında: redundant kesikli + .tow kontrolü
            if (load != null)
                load.Click += (s, e) => Guvenli(f, "YÜKLE sonrası", () =>
                {
                    if (loadSarili && !_sonCizdi) return;   // iptal edildi / çizilmedi
                    var mesaj = new List<string>();
                    string uyari = TowKontrol.Uyari(YolAl(f), OkunanUyeler(f));
                    if (uyari != null) mesaj.Add(uyari);
                    if (chkR.Checked) { mesaj.Add(KesitStili.Uygula(f)); Goruntu.Yenile(); }
                    if (mesaj.Count > 0) Uyari(f, string.Join(" | ", mesaj));
                });

            // ÇALIŞTIR'a ikinci bir olay bağlanır; OutlineDrawing'in kendi çizimi bittikten sonra çalışır.
            if (run != null)
                run.Click += (s, e) => { GorunusYeri.OnbellekAc(); try { Guvenli(f, "ÇALIŞTIR sonrası", () =>
                {
                    Ayarlar.Kaydet(f);
                    if (runSarili && !_sonCizdi)
                    {
                        // Tablo hatası vb.: OutlineDrawing çizmedi. Eski çizime dokunulmaz, sebebi gözden kaçmasın.
                        if (_sonOdMesaji != null)
                            MessageBox.Show(f, "ÇİZİM YENİLENMEDİ.\n\n" + _sonOdMesaji, "Otomatik Kesit", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    var mesaj = new List<string>();     // uyarı etiketine kısa özet
                    var sorun = new List<string>();     // tek uyarı kutusunda gösterilecek sorunlar
                    var sureler = new List<string>();   // komut satırına adım süreleri
                    var toplam = System.Diagnostics.Stopwatch.StartNew();
                    Func<string, Func<string>, string> adim = (ad, is_) =>
                    {
                        var sw = System.Diagnostics.Stopwatch.StartNew();
                        try { return is_(); }
                        finally { sureler.Add(string.Format(CultureInfo.InvariantCulture, "{0} {1:0.00} s", ad, sw.Elapsed.TotalSeconds)); }
                    };
                    if (runSarili && _sonOdMesaji != null) mesaj.Add(_sonOdMesaji.StartsWith("UYARI: ") ? _sonOdMesaji.Substring(7) : _sonOdMesaji);
                    string denetim = adim("denetim", () =>
                    {
                        try { return KesitDenetimi.Yap(f, chkI.Checked); }
                        catch (System.Exception ex) { Komutlar.Yaz("\nKesit denetimi yapılamadı: " + IcHata(ex).Message + "\n"); return null; }
                    });
                    if (denetim != null) { mesaj.Add(denetim); sorun.Add(denetim); }
                    string towUyari = adim(".tow kontrol", () => TowKontrol.Uyari(YolAl(f), OkunanUyeler(f)));
                    if (towUyari != null) mesaj.Add(towUyari);
                    if (chkI.Checked)
                    {
                        string r = adim("iç eleman", () =>
                        {
                            try { return IcElemanlar.Ciz(f, chkR.Checked); }
                            catch (System.Exception ex) { sorun.Add("İç eleman çizgileri çizilemedi: " + IcHata(ex).Message); return null; }
                        });
                        if (r != null) mesaj.Add(r);
                        if (IcElemanlar.SonUyari != null) sorun.Add(IcElemanlar.SonUyari);
                    }
                    else
                    {
                        mesaj.Add("İç eleman: kapalı");
                        sorun.Add("'İç eleman' kutusu kapalı: görünüşlere iç eleman izdüşümleri (tepe, travers, iç çaprazlar) çizilmedi.");
                    }
                    {
                        string r = adim("aile uzatmaları", () =>
                        {
                            try { return UzatmaCizimleri.Ciz(f, chkR.Checked); }
                            catch (System.Exception ex) { sorun.Add("Uzatma çizimleri çizilemedi: " + IcHata(ex).Message); return null; }
                        });
                        if (r != null) mesaj.Add(r);
                    }
                    if (chkR.Checked)
                    {
                        string r = adim("redundant stil", () =>
                        {
                            try { return KesitStili.Uygula(f); }
                            catch (System.Exception ex) { sorun.Add("Redundant çizgi tipi uygulanamadı: " + IcHata(ex).Message); return null; }
                        });
                        if (r != null) mesaj.Add(r);
                    }
                    {
                        string r = adim("kesit yönü", () =>
                        {
                            try { return KesitYonu.Dondur(f); }
                            catch (System.Exception ex) { sorun.Add("Kesitler döndürülemedi: " + IcHata(ex).Message); return null; }
                        });
                        if (r != null) mesaj.Add(r);
                    }
                    if (chk.Checked)
                    {
                        string r = adim("ölçü", () =>
                        {
                            try { return KesitOlculeri.Ekle(); }
                            catch (System.Exception ex) { sorun.Add("Ölçü eklenemedi: " + IcHata(ex).Message); return null; }
                        });
                        if (r != null) mesaj.Add(r);
                    }
                    {
                        string r = adim("kesit işaretleri", () =>
                        {
                            try { return KesitIsaretleri.Ciz(f); }
                            catch (System.Exception ex) { Komutlar.Yaz("\nKesit işaretleri yenilenemedi (OutlineDrawing okları kaldı): " + IcHata(ex).Message + "\n"); return null; }
                        });
                        if (r != null) Komutlar.Yaz("\n" + r + ".\n");
                    }
                    if (chkD.Checked)
                    {
                        string r = adim("kesit dizme", () =>
                        {
                            try { return KesitDizici.Diz(); }
                            catch (System.Exception ex) { sorun.Add("Kesitler dizilemedi (yerleri değişmedi): " + IcHata(ex).Message); return null; }
                        });
                        if (r != null) mesaj.Add(r);
                    }
                    if (chkI.Checked || chkR.Checked || chk.Checked || chkD.Checked) adim("REGEN", () => { Goruntu.Yenile(); return null; });
                    string kapali = KatmanDurumu.Kapalilar(chkI.Checked, chk.Checked);
                    if (kapali != null) sorun.Add(kapali);
                    // Görünüş / kesit düzlemi (XZ) UCS olsun: ölçü ve yazılar tutamaçla düzenlenebilsin
                    Koordinat.Kesit();
                    string sureOzet = string.Format(CultureInfo.InvariantCulture, "Süre: OutlineDrawing {0:0.0} s + eklenti {1:0.0} s",
                        Math.Max(0, _sonCizimSuresi), toplam.Elapsed.TotalSeconds);
                    Komutlar.Yaz("\n" + sureOzet + " (" + string.Join(", ", sureler) + ").\n");
                    mesaj.Add(sureOzet);
                    Uyari(f, string.Join(" | ", mesaj));
                    if (sorun.Count > 0)
                        MessageBox.Show(f, string.Join("\n\n", sorun) + "\n\n(Ayrıntı komut satırında, F2.)", "Otomatik Kesit - kontrol edin",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }); } finally { GorunusYeri.OnbellekKapat(); } };
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

        /// <summary>OutlineDrawing'in .tow'dan okuduğu eleman anahtarları (okunamazsa null).</summary>
        internal static ICollection<string> OkunanUyeler(Form f)
        {
            try
            {
                string yol = YolAl(f);
                if (yol.Length == 0 || !File.Exists(yol)) return null;
                KuleVerisi k = KuleOkuyucu.Oku(f.GetType().Assembly, yol);
                return k != null ? k.Uyeler.Keys : null;
            }
            catch { return null; }
        }

        private static bool _harfYenileniyor;

        /// <summary>Verisi olan kesit satırlarına sırayla A, B, C ... (OutlineDrawing'in IndexToLetters'ı: Z'den sonra AA ...);
        /// içi tamamen boş satırın harfi silinir (OutlineDrawing harfli boş satırı da "dolu" sayıyor).</summary>
        internal static void HarfleriYenile(Form f)
        {
            var grid = Alan(f, "dataGridView1") as DataGridView;
            if (grid == null || !grid.Columns.Contains("ColumnSection") || _harfYenileniyor) return;
            MethodInfo harf = f.GetType().GetMethod("IndexToLetters", BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic,
                null, new[] { typeof(int) }, null);
            _harfYenileniyor = true;
            try
            {
                int i = 0;
                foreach (DataGridViewRow row in grid.Rows)
                {
                    if (row.IsNewRow) continue;
                    bool dolu = row.Cells.Cast<DataGridViewCell>().Any(c => c.OwningColumn.Name != "ColumnSection" &&
                                                                            c.Value != null && c.Value.ToString().Trim().Length > 0);
                    string yeni = "";
                    if (dolu)
                    {
                        yeni = harf != null ? Convert.ToString(harf.Invoke(harf.IsStatic ? null : f, new object[] { i }), CultureInfo.InvariantCulture)
                                            : HarfAdi(i);
                        i++;
                    }
                    object eski = row.Cells["ColumnSection"].Value;
                    if (!string.Equals(eski == null ? "" : eski.ToString(), yeni, StringComparison.Ordinal))
                        row.Cells["ColumnSection"].Value = yeni;
                }
            }
            finally { _harfYenileniyor = false; }
        }

        private static string HarfAdi(int i)
        {
            string s = "";
            for (i++; i > 0; i = (i - 1) / 26) s = (char)('A' + (i - 1) % 26) + s;
            return s;
        }

        internal static string YolAl(Form f)
        {
            var txt = Alan(f, "txtTowerPath") as TextBox;
            return txt != null ? (txt.Text ?? "").Trim() : "";
        }

        private static bool _silmeOnaylandi;

        /// <summary>Butonun mevcut (OutlineDrawing) Click olayını, önce silme kontrolü yapan bir olayın içine alır.
        /// OutlineDrawing'in runtowerfile / loadtowerfile'ı model alanındaki tüm nesneleri siliyor.</summary>
        /// <returns>true: sarmalandı (çizip çizmediği _sonCizdi'de); false: olaylar okunamadı, OutlineDrawing olduğu gibi</returns>
        /// <summary>Butonun OutlineDrawing'in bağladığı Click olayını listeden çıkarıp döndürür (okunamazsa null; o zaman
        /// buton olduğu gibi kalır).</summary>
        private static EventHandler AsilOlayiAl(Button btn)
        {
            if (btn == null) return null;
            try
            {
                PropertyInfo ep = typeof(System.ComponentModel.Component).GetProperty("Events", BindingFlags.Instance | BindingFlags.NonPublic);
                FieldInfo kf = typeof(Control).GetField("EventClick", BindingFlags.Static | BindingFlags.NonPublic)
                            ?? typeof(Control).GetField("ClickEvent", BindingFlags.Static | BindingFlags.NonPublic);   // Mono
                var liste = ep != null ? ep.GetValue(btn, null) as System.ComponentModel.EventHandlerList : null;
                object anahtar = kf != null ? kf.GetValue(null) : null;
                if (liste == null || anahtar == null) return null;
                var asil = liste[anahtar] as EventHandler;
                if (asil == null) return null;
                liste.RemoveHandler(anahtar, asil);
                return asil;
            }
            catch { return null; }
        }

        /// <summary>OutlineDrawing'in Nokta / Yükseklik seçimi Editor.GetPoint sonucunu UCS'den dünyaya çevirmeden kullanıyor:
        /// seçimden önce UCS dünya yapılır (ÇALIŞTIR sonrası UCS kesit düzleminde kalmış olabilir).</summary>
        private static void DunyaUcsIleSec(Button btn)
        {
            if (btn == null) return;
            EventHandler asil = AsilOlayiAl(btn);
            if (asil != null) btn.Click += (s, e) => { Koordinat.Dunya(); asil(s, e); };
            else btn.MouseDown += (s, e) => Koordinat.Dunya();   // olay listesi okunamazsa: fareye basınca (Click'ten önce)
        }

        private static bool SilmeKorumasiEkle(Form f, Button btn, string ad)
        {
            if (btn == null) return false;
            try
            {
                EventHandler asil = AsilOlayiAl(btn);
                if (asil == null) return false;
                btn.Click += (s, e) =>
                {
                    _sonCizdi = false; _sonOdMesaji = null; _sonCizimSuresi = -1;
                    bool devam = true;
                    try { devam = SilmeOnayi(f, ad); } catch { devam = true; }
                    if (devam) devam = OkunabiliyorMu(f, ad);
                    if (!devam) return;
                    long once = Tohum();
                    HataEtiketi(f, "");
                    KatmaniSifirla();
                    Koordinat.Dunya();   // OutlineDrawing dünya koordinatıyla çalışır (nokta seçimi, 3D model)
                    var sure = System.Diagnostics.Stopwatch.StartNew();
                    asil(s, e);
                    _sonCizimSuresi = sure.Elapsed.TotalSeconds;
                    // OutlineDrawing gerçekten çizdi mi? YÜKLE / ÇALIŞTIR uyarı etiketine yalnız çizmeden döndüğünde yazar
                    // (dosya / seçenek yok, kesit tablosunda hata, istisna); ayrıca çizim yoksa yeni nesne eklenmez.
                    long sonra = Tohum();
                    string mesaj = HataEtiketi(f, null);
                    if (!string.IsNullOrWhiteSpace(mesaj)) _sonOdMesaji = mesaj.Trim();
                    _sonCizdi = _sonOdMesaji == null && (once < 0 || sonra < 0 || sonra != once);
                };
                return true;
            }
            catch { return false; /* olaylar okunamazsa koruma eklenmez, OutlineDrawing olduğu gibi çalışır */ }
        }

        private static bool _sonCizdi;       // son YÜKLE / ÇALIŞTIR'da OutlineDrawing çizim yaptı mı
        private static double _sonCizimSuresi = -1;   // OutlineDrawing'in kendi çizim süresi (s)
        private static string _sonOdMesaji;  // OutlineDrawing'in o tıklamada uyarı etiketine yazdığı mesaj

        /// <summary>Veritabanının sonraki tutamaç değeri: yeni nesne eklenince artar (okunamazsa -1).</summary>
        private static long Tohum()
        {
            try
            {
                Document doc = AcApp.DocumentManager.MdiActiveDocument;
                return doc != null ? doc.Database.Handseed.Value : -1;
            }
            catch { return -1; }
        }

        /// <summary>OutlineDrawing'in uyarı etiketi (lblError); yeni != null ise önce o metin yazılır.</summary>
        private static string HataEtiketi(Form f, string yeni)
        {
            try
            {
                var l = Alan(f, "lblError") as Control;
                if (l == null) return null;
                if (yeni != null) l.Text = yeni;
                return l.Text;
            }
            catch { return null; }
        }

        /// <summary>Geçerli katman eklentinin katmanlarından biriyse "0" yapılır: OutlineDrawing kesit / görünüş başlıklarını
        /// geçerli katmana yazar; o katman kapatılınca başlıklar da kaybolmasın.</summary>
        private static void KatmaniSifirla()
        {
            try
            {
                Document doc = AcApp.DocumentManager.MdiActiveDocument;
                if (doc == null) return;
                string c = Convert.ToString(AcApp.GetSystemVariable("CLAYER"), CultureInfo.InvariantCulture);
                if (IcElemanlar.BizimKatman(c) || string.Equals(c, KesitOlculeri.Katman, StringComparison.OrdinalIgnoreCase))
                    using (doc.LockDocument()) { doc.Database.Clayer = doc.Database.LayerZero; }   // modeless formdan: kilit şart
            }
            catch { }
        }

        /// <summary>OutlineDrawing .tow'u okuyamazsa (Run null döner) kendi kodu model alanını sildikten sonra çöküyor ve
        /// hiçbir şey çizilmiyordu: önce okunup okunamadığına bakılır, okunamıyorsa sebep gösterilir, çizime geçilmez.</summary>
        private static bool OkunabiliyorMu(Form f, string ad)
        {
            string hata;
            try
            {
                string yol = YolAl(f);
                if (yol.Length == 0 || !File.Exists(yol)) return true;   // OutlineDrawing kendi uyarısını verir
                if (!KuleOkuyucu.Okunamiyor(f.GetType().Assembly, yol, out hata)) return true;
            }
            catch { return true; }
            string m = "OutlineDrawing bu .tow dosyasını okuyamadı; " + ad + " yapılmadı (model alanı silinmedi).\n\nSebep: " + hata +
                       "\n\nDosyayı PLS-TOWER'da açıp yeniden kaydedin; düzelmezse dosyayı geliştiriciye gönderin.";
            Komutlar.Yaz("\n" + m.Replace("\n\n", " ") + "\n");
            try { MessageBox.Show(f, m, "Otomatik Kesit", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { }
            return false;
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

        /// <summary>chkFrontFace ("Ön Yüz") / chkSideFace: hangi görünüşler çizilecek (kutucuk yoksa ikisi de).</summary>
        internal static void GorunusSecimi(Form f, out bool on, out bool yan)
        {
            var a = Alan(f, "chkFrontFace") as CheckBox;
            var b = Alan(f, "chkSideFace") as CheckBox;
            if (a == null && b == null) { on = yan = true; return; }
            on = a == null || a.Checked;
            yan = b == null || b.Checked;
        }

        private static string KutuYazisi(Form f, string ad, string varsayilan)
        {
            var c = Alan(f, ad) as CheckBox;
            return c != null && !string.IsNullOrWhiteSpace(c.Text) ? c.Text.Trim() : varsayilan;
        }

        private static void GorunusIsaretle(Form f)
        {
            foreach (string ad in new[] { "chkFrontFace", "chkSideFace" })
            {
                var c = Alan(f, ad) as CheckBox;
                if (c != null) c.Checked = true;
            }
        }

        /// <summary>Seçili görünüşlerin eleman anahtarları (KuleVerisi.GorunenAnahtarlar ikisinin birleşimi).</summary>
        internal static ICollection<string> GorunenAnahtarlar(KuleVerisi kule, bool on, bool yan)
        {
            if (on && yan) return kule.GorunenAnahtarlar;
            var l = new HashSet<string>();
            if (on && kule.OnYuz != null) foreach (var c in kule.OnYuz) if (c.Anahtar != null) l.Add(c.Anahtar);
            if (yan && kule.YanYuz != null) foreach (var c in kule.YanYuz) if (c.Anahtar != null) l.Add(c.Anahtar);
            return l;
        }

        private static bool _sonDoldurmaTamam;   // son Oto Kesit tabloyu hazırladı mı (Hepsini Çiz için)
        private static string _otoHesapAnahtar;  // son hesaplanan otomatik kesitler (aynı anahtar: yeniden hesaplanmaz)
        private static List<AutoSectionCandidate> _otoHesapBulunan;
        private static AutoSectionReport _otoHesapRapor;
        private static string _otoTabloAnahtari; // tablo Oto Kesit'ten beri değişmediyse onun anahtarı, yoksa null
        private static bool _tabloYaziliyor;     // tabloyu eklenti yazıyor (kullanıcı değişikliği sayılmaz)

        /// <summary>.tow (yol + zaman + boyut) ve görünüş seçimi; okunamazsa null.</summary>
        internal const string TumElemanKutusuAdi = "chkTumEleman";

        /// <summary>"Tüm elemanlar" işaretli mi (her gizli eleman bir kesitte); işaretsizse sade kesit.</summary>
        internal static bool TumElemanlar(Form f)
        {
            var c = f.Controls.Find(TumElemanKutusuAdi, true).FirstOrDefault() as CheckBox;
            return c != null && c.Checked;
        }

        internal static string OtoAnahtar(string path, bool on, bool yan, bool tum)
        {
            try
            {
                var fi = new FileInfo(path);
                return fi.Exists ? fi.FullName + "|" + fi.LastWriteTimeUtc.Ticks + "|" + fi.Length + "|" + on + "|" + yan + "|" + tum : null;
            }
            catch { return null; }
        }

        /// <summary>Tablo, şu anki .tow ve görünüş seçimi için Oto Kesit'in yazdığı haliyle mi duruyor?</summary>
        internal static bool TabloOtoHalinde(Form f)
        {
            if (_otoTabloAnahtari == null) return false;
            bool on, yan;
            GorunusSecimi(f, out on, out yan);
            return _otoTabloAnahtari == OtoAnahtar(YolAl(f), on, yan, TumElemanlar(f));
        }

        internal static void TabloyuDoldur(Form f)
        {
            _sonDoldurmaTamam = false;
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
                // Yalnız çizilecek görünüşlerdeki elemanlar "görünür": ikisi de işaretsizse ikisi de işaretlenir
                // (OutlineDrawing kutucukları kapalı açıyor; işaretsiz görünüş çizilmez, elemanları hiçbir yerde olmazdı).
                bool onSecili, yanSecili;
                GorunusSecimi(f, out onSecili, out yanSecili);
                bool otoIsaret = !onSecili && !yanSecili;
                if (otoIsaret) { GorunusIsaretle(f); onSecili = yanSecili = true; }
                // Kule dış yüzleri (ön ve yan yüz düzlemi) kesit konusu değildir: yan görünüş çizilmese bile yan yüz bir
                // görünüştür, kesitle (yüz büyüklüğünde "SECTION") verilmez. Görünür sayılan elemanlar her zaman iki yüzün
                // elemanlarıdır; görünüş seçimi yalnız sade modu belirler.
                ICollection<string> gorunen = GorunenAnahtarlar(kule, true, true);
                bool tum = TumElemanlar(f);
                string anahtar = OtoAnahtar(path, onSecili, yanSecili, tum);
                if (anahtar != null && anahtar == _otoHesapAnahtar)
                {
                    bulunan = _otoHesapBulunan; rapor = _otoHesapRapor;   // aynı .tow ve görünüş seçimi: yeniden hesaplanmaz
                }
                else
                {
                    if (gorunen != null && gorunen.Count > 0)
                        bulunan = OtoKesitBulucu.Bul(kule.Uyeler, gorunen, out rapor, !tum);
                    else
                        bulunan = AutoSectionDetector.Detect(kule.Uyeler, new AutoSectionOptions());
                    _otoHesapAnahtar = anahtar; _otoHesapBulunan = bulunan; _otoHesapRapor = rapor;
                }

                if (bulunan.Count == 0)
                {
                    _sonDoldurmaTamam = rapor != null && rapor.HiddenMembers == 0;
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
                    _tabloYaziliyor = true;
                    try { if (cevap == DialogResult.Yes) grid.Rows.Clear(); } finally { _tabloYaziliyor = false; }
                }
                var satirlar = new List<string>();
                _tabloYaziliyor = true;
                try {
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

                HarfleriYenile(f);
                } finally { _tabloYaziliyor = false; }
                _otoTabloAnahtari = OtoAnahtar(path, onSecili, yanSecili, tum);   // tablo bu .tow + seçim için Oto Kesit'in ürettiği hal

                string ozet = string.Format("{0} kesit eklendi", bulunan.Count);
                string onAd = KutuYazisi(f, "chkFrontFace", "Ön Yüz"), yanAd = KutuYazisi(f, "chkSideFace", "Yan Yüz");
                if (otoIsaret) ozet = string.Format("'{0}' ve '{1}' işaretlendi; {2}", onAd, yanAd, ozet);
                else if (!onSecili || !yanSecili) ozet = string.Format("Yalnız '{0}' çizilecek: {1}", onSecili ? onAd : yanAd, ozet);
                if (rapor != null)
                    ozet += string.Format(" (görünmeyen {0}/{1} eleman kapsandı{2})", rapor.CoveredHidden, rapor.HiddenMembers,
                        rapor.UncoveredKeys.Count > 0 ? ", liste komut satırında" : "");
                string towUyari = TowKontrol.Uyari(path, kule.Uyeler.Keys);
                Uyari(f, ozet + ". Kontrol edip ÇALIŞTIR'a basın." + (towUyari != null ? " " + towUyari : ""));

                string log = "\nOtomatik kesitler:\n  " + string.Join("\n  ", satirlar) + "\n";
                if (rapor != null)
                {
                    log += string.Format("Ön/yan görünüşte görünmeyen eleman: {0}, kesitlerle kapsanan: {1}, kapsanmayan: {2}\n",
                        rapor.HiddenMembers, rapor.CoveredHidden, rapor.UncoveredKeys.Count);
                    foreach (string d in rapor.UncoveredDescriptions) log += "  KAPSANMAYAN: " + d + "\n";
                }
                Komutlar.Yaz(log);
                _sonDoldurmaTamam = true;
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

    /// <summary>
    /// Çizgi tipi değişen çizgilerin kesikli görünmesi için REGEN (stil uygulandıktan sonra) ve yakınlaştırma izleyicisinin
    /// başlatılması (bkz. KesikIzleyici).
    /// </summary>
    internal static class Goruntu
    {
        internal static void Yenile()
        {
            Document doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            KesikIzleyici.Etkinlestir(doc);
            if (!KesikIzleyici.Regen(doc))
            {
                try { doc.SendStringToExecute("_.REGEN ", true, false, false); }
                catch { Komutlar.Yaz("\nGörüntü yenilenemedi; kesikli çizgiler için RE (REGEN) yazın.\n"); }
            }
        }
    }

    /// <summary>
    /// AutoCAD kesik deseni REGEN anındaki yakınlaştırmaya göre üretir: desen o an ekranda çok küçükse (ÇALIŞTIR'dan sonra
    /// bütün çizim ekrandayken 150/75 mm birkaç piksel) çizgiyi düz çizer ve yakınlaştırınca, REGEN olmadan, düz kalır.
    /// Kullanıcı LTSCALE'i değiştirince (REGEN yapar) kesikler görünüyordu. Bu izleyici, stil uygulanmış çizimde görünüm
    /// son REGEN'dekinden belirgin biçimde yakınlaşıp durduğunda (VIEWSIZE 1/Oran'dan küçük, iki tik aynı) REGEN yapar.
    /// Uzaklaşınca yalnız taban güncellenir; komut çalışırken (CMDACTIVE) bir şey yapılmaz.
    /// </summary>
    /// <summary>Form ayarlarını (son .tow, görünüş / metin seçenekleri, eklenti kutucukları) %APPDATA%\OtomatikKesit\ayarlar.txt'de
    /// saklar. Okunamazsa / yazılamazsa sessizce geçer.</summary>
    internal static class Ayarlar
    {
        /// <summary>Kullanıcının seçtiği pafta şablonu (null: gömülü ŞA-RA A1).</summary>
        internal static string PaftaSablonu;

        private static string Dosya
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OtomatikKesit", "ayarlar.txt"); }
        }

        private static IEnumerable<FieldInfo> Alanlar(Form f, Type tip)
        {
            return f.GetType().GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                    .Where(x => tip.IsAssignableFrom(x.FieldType));
        }

        internal static void Kaydet(Form f)
        {
            try
            {
                var satir = new List<string>();
                var txt = FormBaglayici.Alan(f, "txtTowerPath") as TextBox;
                if (txt != null) satir.Add("tow=" + (txt.Text ?? "").Trim());
                foreach (FieldInfo fi in Alanlar(f, typeof(CheckBox)))
                {
                    var c = fi.GetValue(f) as CheckBox;
                    if (c != null) satir.Add("kutu:" + fi.Name + "=" + (c.Checked ? "1" : "0"));
                }
                foreach (FieldInfo fi in Alanlar(f, typeof(RadioButton)))
                {
                    var r = fi.GetValue(f) as RadioButton;
                    if (r != null && r.Checked) satir.Add("secim:" + fi.Name + "=1");
                }
                foreach (string ad in new[] { "chkAutoDim", "chkRedundant", "chkIcEleman", "chkKesitDiz", FormBaglayici.TumElemanKutusuAdi })
                {
                    var c = f.Controls.Find(ad, true).FirstOrDefault() as CheckBox;
                    if (c != null) satir.Add("eklenti:" + ad + "=" + (c.Checked ? "1" : "0"));
                }
                if (!string.IsNullOrEmpty(PaftaSablonu)) satir.Add("pafta=" + PaftaSablonu);
                Directory.CreateDirectory(Path.GetDirectoryName(Dosya));
                File.WriteAllLines(Dosya, satir, System.Text.Encoding.UTF8);
            }
            catch { }
        }

        internal static void Yukle(Form f)
        {
            try
            {
                if (!File.Exists(Dosya)) return;
                foreach (string s in File.ReadAllLines(Dosya, System.Text.Encoding.UTF8))
                {
                    int e = s.IndexOf('=');
                    if (e <= 0) continue;
                    string k = s.Substring(0, e), v = s.Substring(e + 1);
                    if (k == "tow")
                    {
                        var txt = FormBaglayici.Alan(f, "txtTowerPath") as TextBox;
                        if (txt != null && (txt.Text ?? "").Trim().Length == 0 && v.Length > 0 && File.Exists(v))
                        {
                            txt.Text = v;
                            PropertyInfo sp = f.GetType().GetProperty("SelectedFilePath", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                            if (sp != null && sp.CanWrite && sp.PropertyType == typeof(string)) sp.SetValue(f, v, null);
                        }
                    }
                    else if (k == "pafta") PaftaSablonu = v;
                    else if (k.StartsWith("kutu:"))
                    {
                        var c = FormBaglayici.Alan(f, k.Substring(5)) as CheckBox;
                        if (c != null) c.Checked = v == "1";
                    }
                    else if (k.StartsWith("secim:"))
                    {
                        var r = FormBaglayici.Alan(f, k.Substring(6)) as RadioButton;
                        if (r != null) r.Checked = true;
                    }
                    else if (k.StartsWith("eklenti:"))
                    {
                        var c = f.Controls.Find(k.Substring(8), true).FirstOrDefault() as CheckBox;
                        // İç eleman kapalı hatırlanmaz: unutulursa görünüşler sessizce eksik kalır
                        if (c != null && !(c.Name == "chkIcEleman" && v == "0")) c.Checked = v == "1";
                    }
                }
            }
            catch { }
        }
    }

    /// <summary>Nesneyi açmadan sınıfını kontrol eder (model alanı taramalarında yalnız ilgili nesneler açılır). Sınıf
    /// okunamazsa true (nesne açılıp eski yoldan denetlenir).</summary>
    internal static class Sinif
    {
        internal static bool Cizgi(ObjectId id) { return Ad(id, "LINE"); }
        internal static bool Yazi(ObjectId id) { return Ad(id, "TEXT"); }

        private static bool Ad(ObjectId id, string dxf)
        {
            try
            {
                var c = id.ObjectClass;
                return c == null || string.Equals(c.DxfName, dxf, StringComparison.OrdinalIgnoreCase);
            }
            catch { return true; }
        }
    }

    /// <summary>ÇALIŞTIR sonrası kesitleri ızgaraya dizer (KesitDizimi). Her kesitin bütün nesneleri (çizgiler, eleman etiketleri,
    /// başlık, ölçüler) birlikte taşınır ve kesit adıyla işaretlenir (XData OTO_KESIT); sonraki ölçü / stil adımları kesitleri
    /// bu işaretle tanır. Tek işlemde: hata olursa hiçbir şey taşınmaz.</summary>
    internal static class KesitDizici
    {
        internal const string Uygulama = "OTO_KESIT";
        private static readonly Regex KesitYazisi = new Regex(@"^SECTION\s+(\S+)$", RegexOptions.IgnoreCase);
        /// <summary>Aile dosyasındaki uzatma görünüşlerinin başlığı ("+9 BODY EXTENSION"): kesit gibi gruplanır ve dizilir.</summary>
        internal static readonly Regex UzatmaYazisi = new Regex(@"^\+?\d+(\.\d+)?\s+(BODY|LEG)\s+EXTENSION$", RegexOptions.IgnoreCase);

        internal static string Etiket(Entity e)
        {
            try
            {
                using (ResultBuffer rb = e.GetXDataForApplication(Uygulama))
                {
                    if (rb == null) return null;
                    foreach (TypedValue tv in rb)
                        if (tv.TypeCode == (int)DxfCode.ExtendedDataAsciiString) return tv.Value as string;
                }
            }
            catch { }
            return null;
        }

        internal class Baslik { public string Ad; public double U, V, W; public bool XZ; public ObjectId Id; }

        /// <summary>Model alanındaki kesitler (başlık sırasıyla) ve görünüş başlıkları.</summary>
        internal class Gruplar
        {
            public bool XZ;
            public List<Baslik> Kesitler = new List<Baslik>();
            public Dictionary<string, List<ObjectId>> Nesneler = new Dictionary<string, List<ObjectId>>();
            public Dictionary<string, double[]> Kutular = new Dictionary<string, double[]>();
            public Baslik On, Yan;
        }

        // { u0, v0, u1, v1 } (XZ düzleminde v = Z, XY'de v = Y); okunamazsa null
        internal static double[] Kutu(Entity e, bool xz)
        {
            Point3d a, b;
            try { Extents3d x = e.GeometricExtents; a = x.MinPoint; b = x.MaxPoint; }
            catch
            {
                var t = e as DBText; var l = e as Line;
                if (t != null) { a = b = t.Position; }
                else if (l != null) { a = l.StartPoint; b = l.EndPoint; }
                else return null;
            }
            double va = xz ? a.Z : a.Y, vb = xz ? b.Z : b.Y;
            return new[] { Math.Min(a.X, b.X), Math.Min(va, vb), Math.Max(a.X, b.X), Math.Max(va, vb) };
        }

        internal static double[] Birlestir(double[] a, double[] b)
        {
            return a == null ? b : b == null ? a : new[] { Math.Min(a[0], b[0]), Math.Min(a[1], b[1]), Math.Max(a[2], b[2]), Math.Max(a[3], b[3]) };
        }

        /// <summary>Kesit başlıkları ve her kesitin nesneleri. Dizilmiş kesitler XData'dan; dizilmemişler (OutlineDrawing'in tek sırası)
        /// en yakın başlığa, kesit aralığının yarısı içinde ve bir görünüş başlığından daha yakınsa. haric: bu katmanlar atlanır.</summary>
        internal static Gruplar Grupla(Transaction tr, BlockTableRecord ms, Func<Entity, bool> haric = null)
        {
            var g = new Gruplar();
            var basliklar = new List<Baslik>();
            var gorunus = new List<Baslik>();
            foreach (ObjectId id in ms)
            {
                if (!Sinif.Yazi(id)) continue;
                var t = tr.GetObject(id, OpenMode.ForRead) as DBText;
                if (t == null || t.TextString == null) continue;
                string metin = t.TextString.Trim();
                Point3d p = t.HorizontalMode == TextHorizontalMode.TextLeft && t.VerticalMode == TextVerticalMode.TextBase ? t.Position : t.AlignmentPoint;
                bool xz = Math.Abs(t.Normal.Y) > 0.9, xy = Math.Abs(t.Normal.Z) > 0.9;
                if (!xz && !xy) continue;
                var b = new Baslik { U = p.X, V = xz ? p.Z : p.Y, W = xz ? p.Y : p.Z, XZ = xz, Id = id };
                if (metin.Equals(GorunusYeri.OnBaslik, StringComparison.OrdinalIgnoreCase)) { b.Ad = GorunusYeri.OnBaslik; gorunus.Add(b); if (g.On == null) g.On = b; continue; }
                if (metin.Equals(GorunusYeri.YanBaslik, StringComparison.OrdinalIgnoreCase)) { b.Ad = GorunusYeri.YanBaslik; gorunus.Add(b); if (g.Yan == null) g.Yan = b; continue; }
                if (t.Height < 250) continue;
                Match m = KesitYazisi.Match(metin);
                if (m.Success && !basliklar.Any(o => o.Ad == m.Groups[1].Value)) { b.Ad = m.Groups[1].Value; basliklar.Add(b); }
                else if (!m.Success && UzatmaYazisi.IsMatch(metin) && !basliklar.Any(o => o.Ad == metin)) { b.Ad = metin; basliklar.Add(b); }
            }
            int nXZ = basliklar.Count(b => b.XZ) + gorunus.Count(b => b.XZ), nXY = basliklar.Count(b => !b.XZ) + gorunus.Count(b => !b.XZ);
            g.XZ = nXZ >= nXY;
            g.Kesitler = basliklar.Where(b => b.XZ == g.XZ).OrderBy(b => b.U).ToList();
            if (g.On != null && g.On.XZ != g.XZ) g.On = null;
            if (g.Yan != null && g.Yan.XZ != g.XZ) g.Yan = null;
            foreach (var b in g.Kesitler) { g.Nesneler[b.Ad] = new List<ObjectId>(); g.Kutular[b.Ad] = null; }
            if (g.Kesitler.Count == 0) return g;

            double enKucuk = double.MaxValue;
            for (int i = 1; i < g.Kesitler.Count; i++)
                if (g.Kesitler[i].U - g.Kesitler[i - 1].U > 1000) enKucuk = Math.Min(enKucuk, g.Kesitler[i].U - g.Kesitler[i - 1].U);
            double yari = enKucuk < double.MaxValue ? enKucuk / 2 : 15000;
            var gorunusU = gorunus.Where(b => b.XZ == g.XZ).Select(b => b.U).ToList();
            foreach (ObjectId id in ms)
            {
                if (id.IsErased) continue;
                var e = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (e == null || (haric != null && haric(e))) continue;
                string ad = Etiket(e);
                double[] k = Kutu(e, g.XZ);
                if (k == null) continue;
                if (ad == null || !g.Nesneler.ContainsKey(ad))
                {
                    double um = (k[0] + k[2]) / 2;
                    Baslik en = g.Kesitler.OrderBy(b => Math.Abs(b.U - um)).First();
                    if (Math.Abs(en.U - um) > yari) continue;
                    if (gorunusU.Any(x => Math.Abs(x - um) < Math.Abs(en.U - um))) continue;   // görünüşe daha yakın
                    if (Etiket(e) != null) continue;   // başka (silinmiş) kesitin işareti
                    ad = en.Ad;
                }
                g.Nesneler[ad].Add(id);
                g.Kutular[ad] = Birlestir(g.Kutular[ad], k);
            }
            g.Kesitler = g.Kesitler.Where(b => g.Kutular[b.Ad] != null).ToList();
            return g;
        }

        internal static void UygulamaKaydi(Database db, Transaction tr)
        {
            var rat = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
            if (rat.Has(Uygulama)) return;
            rat.UpgradeOpen();
            var r = new RegAppTableRecord { Name = Uygulama };
            rat.Add(r);
            tr.AddNewlyCreatedDBObject(r, true);
        }

        /// <summary>Kesitin nesnelerini kaydırır ve kesit adıyla işaretler.</summary>
        internal static int Tasi(Transaction tr, IEnumerable<ObjectId> ids, string ad, Vector3d v)
        {
            int n = 0;
            Matrix3d m = Matrix3d.Displacement(v);
            bool kaymaVar = v.Length > 1e-6;
            foreach (ObjectId id in ids)
            {
                var e = (Entity)tr.GetObject(id, OpenMode.ForWrite);
                if (kaymaVar) e.TransformBy(m);
                if (ad != null)
                    e.XData = new ResultBuffer(new TypedValue((int)DxfCode.ExtendedDataRegAppName, Uygulama),
                                               new TypedValue((int)DxfCode.ExtendedDataAsciiString, ad));
                n++;
            }
            return n;
        }

        internal static string Diz()
        {
            Document doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return null;
            Database db = doc.Database;
            using (doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var ms = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);
                Gruplar g = Grupla(tr, ms, Pafta.PaftaNesnesi);
                if (g.Kesitler.Count < 2) { tr.Commit(); return null; }
                int satir;
                List<double[]> kayma = KesitDizimi.Yerlestir(g.Kesitler.Select(b => g.Kutular[b.Ad]).ToList(), out satir);
                UygulamaKaydi(db, tr);
                int tasinan = 0;
                for (int i = 0; i < g.Kesitler.Count; i++)
                {
                    Vector3d v = g.XZ ? new Vector3d(kayma[i][0], 0, kayma[i][1]) : new Vector3d(kayma[i][0], kayma[i][1], 0);
                    tasinan += Tasi(tr, g.Nesneler[g.Kesitler[i].Ad], g.Kesitler[i].Ad, v);
                }
                tr.Commit();
                Komutlar.Yaz(string.Format("\nKesitler dizildi: {0} kesit, {1} satır ({2} nesne).\n", g.Kesitler.Count, satir, tasinan));
                return string.Format("Kesitler: {0} kesit {1} satırda", g.Kesitler.Count, satir);
            }
        }
    }

    /// <summary>Malzeme listesi (Group summary): .tow grup tablosundan, PLS'teki sırayla. Sütunlar şirket paftasındaki gibi:
    /// Group Label, Group Desc., Angle Size, Strength(Mpa) (malzeme adındaki sayı: S355 -> 355.0).</summary>
    public static class MalzemeListesi
    {
        public static readonly string[] Basliklar = { "Group Label", "Group Desc.", "Angle Size", "Strength(Mpa)" };
        private static readonly Regex Alan = new Regex("'([^']*)'|(\\S+)");

        public static List<string[]> Oku(string towPath)
        {
            var sonuc = new List<string[]>();
            if (string.IsNullOrEmpty(towPath) || !File.Exists(towPath)) return sonuc;
            string[] satirlar = TowYama.Satirlar(towPath);
            for (int i = 0; i < satirlar.Length; i++)
            {
                if (satirlar[i].IndexOf("; group label", StringComparison.OrdinalIgnoreCase) < 0) continue;
                int n;
                if (!int.TryParse(satirlar[i].Substring(0, satirlar[i].IndexOf(';')).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out n)) break;
                for (int k = i + 1; k <= i + n && k < satirlar.Length; k++)
                {
                    var t = Alan.Matches(satirlar[k]).Cast<Match>().Select(m => m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value).ToList();
                    if (t.Count < 4) continue;
                    sonuc.Add(new[] { t[0].Trim(), t[1].Trim(), t[2].Trim(), Dayanim(t[3].Trim()) });
                }
                break;
            }
            return sonuc;
        }

        /// <summary>"S355" -> "355.0"; sayı yoksa malzeme adı.</summary>
        public static string Dayanim(string malzeme)
        {
            Match m = Regex.Match(malzeme ?? "", "(\\d+(?:[.,]\\d+)?)");
            double d;
            if (m.Success && double.TryParse(m.Groups[1].Value.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out d))
                return d.ToString("0.0", CultureInfo.InvariantCulture);
            return malzeme ?? "";
        }
    }

    /// <summary>Pafta yerleşimi (AutoCAD'den bağımsız). Kâğıt koordinatı: şablonun (0,0)-(G,Y) dikdörtgeni; bölgeler şirketin A1
    /// paftasından (YAA_OUTLINE_rev0) ölçülmüş oranlar. Ön görünüş solda, yan görünüş yanında (yalnız yatay kayar), kesitler
    /// kalan alana (sol blok, malzeme listesinin altı) raf raf; sığmazsa ölçek büyür. Ölçek 5'in katına yuvarlanır (1:x).</summary>
    public static class PaftaPlani
    {
        public static double[] Ic = { 0.00842, 0.01191, 0.99158, 0.98809 };      // iç çerçeve
        public static double[] Antet = { 0.7631, 0.01191, 0.99158, 0.3952 };     // sağ alt: antet + genel notlar
        public static double[] Liste = { 0.5755, 0.4389, 0.9845, 0.9754 };       // sağ üst: malzeme listesi
        public static double Bosluk = 6;                                          // kâğıtta öğeler arası (mm)

        public class Sonuc
        {
            public double Olcek;              // model mm / kâğıt mm
            public double OrjinU, OrjinV;     // şablonun (0,0) noktasının model koordinatı
            public double YanDu;              // yan görünüşün yatay kayması (tam mm)
            public List<double[]> KesitKayma = new List<double[]>();   // { du, dv }
            public double[] ListeAlani;       // kâğıtta { x0, y0, x1, y1 } (liste yoksa null)
            public bool Sigdi;
        }

        static double[] Olc(double[] f, double G, double Y) { return new[] { f[0] * G, f[1] * Y, f[2] * G, f[3] * Y }; }

        /// <param name="on">Ön görünüş kutusu (model { u0, v0, u1, v1 }); yerinde kalır, pafta onun etrafına kurulur.</param>
        /// <param name="yan">Yan görünüş kutusu (yoksa null).</param>
        public static Sonuc Hesapla(double[] on, double[] yan, List<double[]> kesitler, bool liste, double G = 841, double Y = 594.65)
        {
            double[] ic = Olc(Ic, G, Y), ant = Olc(Antet, G, Y), lst = Olc(Liste, G, Y);
            double B = Bosluk;
            var L = new[] { ic[0] + B, ic[1] + B, (liste ? lst[0] : ant[0]) - B, ic[3] - B };
            var bolgeler = new List<double[]>();
            if (liste) bolgeler.Add(new[] { lst[0], ic[1] + B, ant[0] - B, lst[1] - B });          // listenin altı, antetin solu
            else bolgeler.Add(new[] { ant[0] + B, ant[3] + B, ic[2] - B, ic[3] - B });            // antetin üstü
            double hV = Math.Max(on[3] - on[1], yan != null ? Math.Max(yan[3], on[3]) - Math.Min(yan[1], on[1]) : 0);
            double s = Math.Max(1, Math.Ceiling(hV / (L[3] - L[1]) / 5) * 5);
            for (int deneme = 0; deneme < 400; deneme++, s += 5)
            {
                var r = Dene(on, yan, kesitler, s, L, bolgeler, B);
                if (r == null) continue;
                r.ListeAlani = liste ? lst : null;
                return r;
            }
            return new Sonuc { Sigdi = false };
        }

        static Sonuc Dene(double[] on, double[] yan, List<double[]> kesitler, double s, double[] L, List<double[]> ekBolgeler, double B)
        {
            double wF = (on[2] - on[0]) / s, hF = (on[3] - on[1]) / s;
            double pxF = L[0], pyF = L[3] - hF;                    // ön görünüş: sol üst
            if (yan != null) pyF = Math.Min(pyF, L[3] - (yan[3] - on[1]) / s);   // yan görünüş de tavana sığsın
            if (pyF < L[1] || pxF + wF > L[2]) return null;
            double sonX = pxF + wF;
            double pxS = 0;
            if (yan != null)
            {
                pxS = sonX + B;
                double wS = (yan[2] - yan[0]) / s, pyS0 = pyF + (yan[1] - on[1]) / s, pyS1 = pyF + (yan[3] - on[1]) / s;
                if (pxS + wS > L[2] || pyS0 < L[1] || pyS1 > L[3]) return null;
                sonX = pxS + wS;
            }
            // kesit bölgeleri: görünüşlerin sağı (görünüş yüksekliğinde), görünüşlerin altı (tam genişlik), ek bölgeler
            double gorunusAlt = pyF;
            if (yan != null) gorunusAlt = Math.Min(gorunusAlt, pyF + (yan[1] - on[1]) / s);
            var bolgeler = new List<double[]>
            {
                new[] { sonX + B, gorunusAlt, L[2], L[3] },
                new[] { L[0], L[1], L[2], gorunusAlt - B }
            };
            bolgeler.AddRange(ekBolgeler);
            bolgeler = bolgeler.Where(b => b[2] - b[0] > 1 && b[3] - b[1] > 1).ToList();
            var r = new Sonuc { Olcek = s, OrjinU = on[0] - pxF * s, OrjinV = on[1] - pyF * s, Sigdi = true };
            if (yan != null) r.YanDu = Math.Round(r.OrjinU + pxS * s - yan[0]);
            // Her bölgede raf (satır) düzeni; kesit, sırayla, sığdığı ilk bölgeye (önce o bölgenin açık satırına, sonra yeni
            // satıra). Geniş bir kesit bir bölgeye sığmayınca küçükler o bölgeyi boş bırakmaz.
            int n = bolgeler.Count;
            var x = new double[n]; var ust = new double[n]; var satirH = new double[n];
            for (int i = 0; i < n; i++) { x[i] = bolgeler[i][0]; ust[i] = bolgeler[i][3]; }
            foreach (double[] k in kesitler)
            {
                double w = (k[2] - k[0]) / s, h = (k[3] - k[1]) / s, px = 0, py = 0;
                bool kondu = false;
                for (int i = 0; i < n && !kondu; i++)
                {
                    double[] R = bolgeler[i];
                    if (w > R[2] - R[0]) continue;
                    if (x[i] + w <= R[2] && ust[i] - h >= R[1]) { px = x[i]; py = ust[i] - h; }                 // açık satır
                    else if (ust[i] - satirH[i] - B - h >= R[1])                                               // yeni satır
                    { ust[i] -= satirH[i] + B; x[i] = R[0]; satirH[i] = 0; px = x[i]; py = ust[i] - h; }
                    else continue;
                    x[i] = px + w + B; satirH[i] = Math.Max(satirH[i], h); kondu = true;
                }
                if (!kondu) return null;
                r.KesitKayma.Add(new[] { r.OrjinU + px * s - k[0], r.OrjinV + py * s - k[1] });
            }
            return r;
        }

        /// <summary>Malzeme listesi tabloları (kâğıt koordinatı): satırlar sütunlara bölünür. Döner: her hücre { x0, y0, x1, y1 } ve
        /// yazı yüksekliği; çizim için.</summary>
        public static int TabloSutunu = 3;
        public static double[] SutunOranlari = { 0.17, 0.45, 0.22, 0.16 };
    }

    /// <summary>Pafta: şablon (gömülü ŞA-RA A1 ya da seçilen DWG/DXF) çerçevesine ön / yan görünüş, kesitler ve malzeme listesi.</summary>
    internal static class Pafta
    {
        internal const string Katman = "OTO_PAFTA", ListeKatmani = "OTO_LISTE";
        private const string GomuluAd = "SARA_A1_pafta.dxf";

        internal static bool PaftaNesnesi(Entity e)
        {
            return string.Equals(e.Layer, Katman, StringComparison.OrdinalIgnoreCase) || string.Equals(e.Layer, ListeKatmani, StringComparison.OrdinalIgnoreCase);
        }

        private static string GomuluSablon()
        {
            string yol = Path.Combine(Path.GetTempPath(), "OtoKesit_" + GomuluAd);
            using (Stream st = typeof(Pafta).Assembly.GetManifestResourceStream(GomuluAd))
            {
                if (st == null) throw new FileNotFoundException("Gömülü pafta şablonu yok: " + GomuluAd);
                using (var fs = File.Create(yol)) st.CopyTo(fs);
            }
            return yol;
        }

        private static ObjectId SablonBlogu(Database db, string sablon, out double G, out double Y)
        {
            string dosya = string.IsNullOrEmpty(sablon) ? null : sablon;
            string ad = "OTO_PAFTA_" + Regex.Replace(Path.GetFileNameWithoutExtension(dosya ?? GomuluAd), "[^A-Za-z0-9_]", "_");
            ObjectId id = ObjectId.Null;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (bt.Has(ad)) id = bt[ad];   // önceki paftadan: yeniden eklenmez
                tr.Commit();
            }
            if (id.IsNull)
            {
                if (dosya == null) dosya = GomuluSablon();
                using (var tdb = new Database(false, true))
                {
                    if (dosya.EndsWith(".dxf", StringComparison.OrdinalIgnoreCase)) tdb.DxfIn(dosya, Path.Combine(Path.GetTempPath(), "OtoKesit_dxfin.log"));
                    else { tdb.ReadDwgFile(dosya, FileOpenMode.OpenForReadAndAllShare, true, ""); tdb.CloseInput(true); }
                    id = db.Insert(ad, tdb, true);
                }
            }
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                double[] k = null;
                foreach (ObjectId e in btr)
                {
                    var ent = tr.GetObject(e, OpenMode.ForRead) as Entity;
                    if (ent == null) continue;
                    try { Extents3d x = ent.GeometricExtents; k = KesitDizici.Birlestir(k, new[] { x.MinPoint.X, x.MinPoint.Y, x.MaxPoint.X, x.MaxPoint.Y }); } catch { }
                }
                tr.Commit();
                if (k == null) throw new InvalidOperationException("Pafta şablonu boş: " + (dosya ?? GomuluAd));
                // Şablonun sol altı (0,0) olmalı; bölge oranları şablonun genişlik / yüksekliğine göre
                G = k[2]; Y = k[3];
            }
            return id;
        }

        private static ObjectId KatmanHazirla(Database db, Transaction tr, string ad, short renk)
        {
            var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (!lt.Has(ad))
            {
                lt.UpgradeOpen();
                var r = new LayerTableRecord { Name = ad, Color = AcColor.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, renk) };
                lt.Add(r); tr.AddNewlyCreatedDBObject(r, true);
            }
            var k = (LayerTableRecord)tr.GetObject(lt[ad], OpenMode.ForRead);
            if (k.IsLocked || k.IsOff || k.IsFrozen) { k.UpgradeOpen(); k.IsLocked = false; k.IsOff = false; k.IsFrozen = false; }
            return lt[ad];
        }

        internal static string Olustur(Form f, string sablon)
        {
            Document doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return "Pafta: açık çizim yok.";
            Database db = doc.Database;
            string yol = f != null ? FormBaglayici.YolAl(f) : "";
            List<string[]> liste = MalzemeListesi.Oku(yol);
            using (doc.LockDocument())
            {
                double G, Y;
                ObjectId blok = SablonBlogu(db, sablon, out G, out Y);
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    var ms = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);
                    // önceki pafta ve liste silinir
                    foreach (ObjectId id in ms)
                    {
                        var e = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (e != null && PaftaNesnesi(e)) { e.UpgradeOpen(); e.Erase(); }
                    }
                    KesitDizici.Gruplar g = KesitDizici.Grupla(tr, ms, PaftaNesnesi);
                    if (g.On == null) { tr.Commit(); return "Pafta: çizimde TRANSVERSE FACE yok (önce ÇALIŞTIR, Ön Yüz işaretli)."; }

                    // görünüşlerin nesneleri: kesit olmayan, en yakın görünüş başlığı
                    var kesitNesnesi = new HashSet<ObjectId>(g.Nesneler.Values.SelectMany(x => x));
                    double sinir = g.Yan != null ? Math.Abs(g.Yan.U - g.On.U) * 0.75 : 60000;
                    var onIds = new List<ObjectId>(); var yanIds = new List<ObjectId>();
                    double[] onK = null, yanK = null;
                    foreach (ObjectId id in ms)
                    {
                        if (id.IsErased || kesitNesnesi.Contains(id)) continue;
                        var e = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (e == null || PaftaNesnesi(e) || KesitDizici.Etiket(e) != null) continue;
                        double[] k = KesitDizici.Kutu(e, g.XZ);
                        if (k == null) continue;
                        double um = (k[0] + k[2]) / 2;
                        double dOn = Math.Abs(um - g.On.U), dYan = g.Yan != null ? Math.Abs(um - g.Yan.U) : double.MaxValue;
                        if (Math.Min(dOn, dYan) > sinir) continue;
                        if (dOn <= dYan) { onIds.Add(id); onK = KesitDizici.Birlestir(onK, k); }
                        else { yanIds.Add(id); yanK = KesitDizici.Birlestir(yanK, k); }
                    }
                    if (onK == null) { tr.Commit(); return "Pafta: ön görünüş bulunamadı."; }

                    var kutular = g.Kesitler.Select(b => g.Kutular[b.Ad]).ToList();
                    PaftaPlani.Sonuc p = PaftaPlani.Hesapla(onK, yanK, kutular, liste.Count > 0, G, Y);
                    if (!p.Sigdi) { tr.Commit(); return "Pafta: yerleşim bulunamadı."; }

                    // kâğıt -> model dönüşümü: ölçek, düzlem (XZ: kâğıt y -> model Z), öteleme
                    double W = g.On.W;
                    Matrix3d duzlem = g.XZ ? Matrix3d.Rotation(Math.PI / 2, Vector3d.XAxis, Point3d.Origin) : Matrix3d.Identity;
                    Vector3d orjin = g.XZ ? new Vector3d(p.OrjinU, W, p.OrjinV) : new Vector3d(p.OrjinU, p.OrjinV, W);
                    Matrix3d M = Matrix3d.Displacement(orjin) * duzlem * Matrix3d.Scaling(p.Olcek, Point3d.Origin);

                    KesitDizici.UygulamaKaydi(db, tr);
                    int tasinan = 0;
                    if (g.Yan != null && Math.Abs(p.YanDu) > 0.5)
                        tasinan += KesitDizici.Tasi(tr, yanIds, null, new Vector3d(p.YanDu, 0, 0));
                    for (int i = 0; i < g.Kesitler.Count; i++)
                    {
                        double[] k = p.KesitKayma[i];
                        Vector3d v = g.XZ ? new Vector3d(k[0], 0, k[1]) : new Vector3d(k[0], k[1], 0);
                        tasinan += KesitDizici.Tasi(tr, g.Nesneler[g.Kesitler[i].Ad], g.Kesitler[i].Ad, v);
                    }

                    // şablon
                    KatmanHazirla(db, tr, Katman, 7);
                    var br = new BlockReference(Point3d.Origin, blok) { Layer = Katman };
                    br.TransformBy(M);
                    ms.AppendEntity(br); tr.AddNewlyCreatedDBObject(br, true);
                    OlcekYaz(tr, blok, p.Olcek);

                    // malzeme listesi
                    if (liste.Count > 0) ListeCiz(db, tr, ms, liste, p.ListeAlani, M);
                    tr.Commit();
                    string ozet = string.Format("Pafta: ölçek 1:{0:0}, {1} kesit, malzeme listesi {2} satır", p.Olcek, g.Kesitler.Count, liste.Count);
                    Komutlar.Yaz("\n" + ozet + " (" + tasinan + " nesne taşındı).\n");
                    return ozet;
                }
            }
        }

        /// <summary>Şablondaki ölçek hücresi ("NTS" ya da önceki "1:x") güncellenir.</summary>
        private static void OlcekYaz(Transaction tr, ObjectId blok, double s)
        {
            var btr = (BlockTableRecord)tr.GetObject(blok, OpenMode.ForRead);
            foreach (ObjectId id in btr)
            {
                var t = tr.GetObject(id, OpenMode.ForRead) as DBText;
                if (t == null || t.TextString == null) continue;
                string m = t.TextString.Trim();
                if (m.Equals("NTS", StringComparison.OrdinalIgnoreCase) || Regex.IsMatch(m, "^1:\\d+$"))
                {
                    t.UpgradeOpen();
                    t.TextString = "1:" + s.ToString("0", CultureInfo.InvariantCulture);
                }
            }
        }

        private static void ListeCiz(Database db, Transaction tr, BlockTableRecord ms, List<string[]> liste, double[] alan, Matrix3d M)
        {
            ObjectId katman = KatmanHazirla(db, tr, ListeKatmani, 7);
            int k = PaftaPlani.TabloSutunu;
            int satirBasi = (int)Math.Ceiling(liste.Count / (double)k);
            double ara = 4, w = (alan[2] - alan[0] - (k - 1) * ara) / k, H = alan[3] - alan[1];
            double rh = Math.Min(6.0, H / (satirBasi + 1)), th = Math.Min(2.5, 0.55 * rh), pad = 0.8;
            double[] oran = PaftaPlani.SutunOranlari;
            Action<Point3d, Point3d> cizgi = (a, b) =>
            {
                var l = new Line(a, b); l.SetDatabaseDefaults(db); l.LayerId = katman; l.TransformBy(M);
                ms.AppendEntity(l); tr.AddNewlyCreatedDBObject(l, true);
            };
            Action<string, double, double, double> yazi = (metin, x0, ym, sw) =>
            {
                if (string.IsNullOrEmpty(metin)) return;
                var t = new DBText(); t.SetDatabaseDefaults(db); t.LayerId = katman;
                t.TextString = metin; t.Height = th;
                double tahmin = metin.Length * 0.62 * th, sigan = sw - 2 * pad;
                if (tahmin > sigan && tahmin > 0) t.WidthFactor = Math.Max(0.45, sigan / tahmin);
                t.HorizontalMode = TextHorizontalMode.TextLeft; t.VerticalMode = TextVerticalMode.TextVerticalMid;
                t.AlignmentPoint = new Point3d(x0 + pad, ym, 0);
                t.TransformBy(M);
                ms.AppendEntity(t); tr.AddNewlyCreatedDBObject(t, true);
                try { t.AdjustAlignment(db); } catch { }
            };
            for (int c = 0; c < k; c++)
            {
                int bas = c * satirBasi, son = Math.Min(liste.Count, bas + satirBasi);
                if (bas >= son) break;
                double x0 = alan[0] + c * (w + ara), y1 = alan[3];
                int satirSay = son - bas + 1;
                double y0 = y1 - satirSay * rh;
                var xs = new List<double> { x0 };
                foreach (double o in oran) xs.Add(xs[xs.Count - 1] + o * w);
                for (int r = 0; r <= satirSay; r++) cizgi(new Point3d(x0, y1 - r * rh, 0), new Point3d(x0 + w, y1 - r * rh, 0));
                foreach (double x in xs) cizgi(new Point3d(x, y0, 0), new Point3d(x, y1, 0));
                for (int j = 0; j < 4; j++) yazi(MalzemeListesi.Basliklar[j], xs[j], y1 - rh / 2, xs[j + 1] - xs[j]);
                for (int r = bas; r < son; r++)
                    for (int j = 0; j < 4; j++) yazi(liste[r][j], xs[j], y1 - (r - bas + 1.5) * rh, xs[j + 1] - xs[j]);
            }
        }
    }

    /// <summary>Eklentinin katmanları kapalı / dondurulmuşsa çizgiler çizilir ama ekranda görünmez; kullanıcı fark etsin.</summary>
    internal static class KatmanDurumu
    {
        /// <summary>Açık olan özelliğin (iç eleman / ölçü) katmanı çizimden sonra hâlâ kapalı ya da dondurulmuşsa uyarı (adımlar
        /// katmanı zaten açar; bu bir güvenlik ağı).</summary>
        internal static string Kapalilar(bool icEleman, bool olcu)
        {
            try
            {
                Document doc = AcApp.DocumentManager.MdiActiveDocument;
                if (doc == null) return null;
                var adlar = new List<string>();
                if (icEleman) { adlar.Add(IcElemanlar.Katman); adlar.Add(IcElemanlar.KatmanUc); }
                if (olcu) adlar.Add(KesitOlculeri.Katman);
                var l = new List<string>();
                using (doc.LockDocument())
                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(doc.Database.LayerTableId, OpenMode.ForRead);
                    foreach (string ad in adlar)
                    {
                        if (!lt.Has(ad)) continue;
                        var k = (LayerTableRecord)tr.GetObject(lt[ad], OpenMode.ForRead);
                        if (k.IsOff || k.IsFrozen)
                            l.Add(ad + (k.IsFrozen ? " (dondurulmuş)" : " (kapalı)"));
                    }
                    tr.Commit();
                }
                return l.Count == 0 ? null
                    : "Şu katman(lar) kapalı, çizgileri / ölçüleri ekranda GÖRÜNMEZ: " + string.Join(", ", l) +
                      ". Görmek için LAYER ile açın (OTO_IC_ELEMAN: iç eleman çizgileri, OTO_KESIT_OLCU: ölçüler).";
            }
            catch { return null; }
        }
    }

    /// <summary>UCS ayarı. OutlineDrawing görünüş ve kesitleri XZ düzlemine (normal -Y) çizer; UCS dünya (XY) iken AutoCAD bu
    /// düzlemdeki ölçü / yazıların tutamaçlarıyla yalnız taşımaya izin veriyordu.</summary>
    internal static class Koordinat
    {
        /// <summary>UCS: X = dünya X, Y = dünya Z (normal -Y), başlangıç (0,0,0).</summary>
        internal static bool Kesit()
        {
            return Ayarla(() => Matrix3d.AlignCoordinateSystem(Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis,
                                                                Point3d.Origin, Vector3d.XAxis, Vector3d.ZAxis, -Vector3d.YAxis));
        }

        internal static bool Dunya() { return Ayarla(() => Matrix3d.Identity); }

        /// <summary>Hiçbir durumda hata fırlatmaz (OutlineDrawing'in butonlarından önce çağrılıyor).</summary>
        private static bool Ayarla(Func<Matrix3d> m)
        {
            try
            {
                Document doc = AcApp.DocumentManager.MdiActiveDocument;
                if (doc == null) return false;
                Matrix3d u = m();
                using (doc.LockDocument()) { doc.Editor.CurrentUserCoordinateSystem = u; }
                return true;
            }
            catch { return false; }
        }
    }

    internal static class KesikIzleyici
    {
        /// <summary>Son REGEN'e göre bu kat yakınlaşınca yeniden üretilir (büyük çizimde her REGEN takılma yapar).</summary>
        public static double Oran = 2.5;
        /// <summary>Yakınlaştırma bittikten sonra bu kadar tik (700 ms) beklenir; tekerlek dönerken REGEN olmaz.</summary>
        public static int BekleTik = 2;
        private static int _sabit;
        private static readonly HashSet<string> _belgeler = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private static string _belge;
        private static double _taban = -1, _onceki = -1;

        internal static void Etkinlestir(Document doc)
        {
            _belgeler.Add(doc.Name);
            _belge = doc.Name;
            _taban = _onceki = GorunumBoyu();
        }

        internal static void Tik()
        {
            Document doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null || !_belgeler.Contains(doc.Name)) return;
            if (Convert.ToInt32(AcApp.GetSystemVariable("CMDACTIVE")) != 0) return;
            double v = GorunumBoyu();
            if (v <= 0) return;
            if (!string.Equals(_belge, doc.Name, StringComparison.OrdinalIgnoreCase)) { _belge = doc.Name; _taban = _onceki = v; return; }
            bool durdu = Math.Abs(v - _onceki) <= 1e-9 * Math.Max(1.0, v);
            _onceki = v;
            if (!durdu) { _sabit = 0; return; }       // tekerlekle yakınlaştırma sürüyor
            if (++_sabit < BekleTik) return;
            if (v > _taban) { _taban = v; return; }   // uzaklaştı: bir sonraki yakınlaşmada yeniden üret
            if (v < _taban / Oran)
            {
                _taban = v;
                if (!Regen(doc))
                {
                    try { doc.SendStringToExecute("_.REGEN ", false, false, false); } catch { }
                }
            }
        }

        /// <returns>REGEN yapılabildiyse true</returns>
        internal static bool Regen(Document doc)
        {
            try
            {
                using (doc.LockDocument()) { doc.Editor.Regen(); }
                return true;
            }
            catch { return false; }
        }

        private static double GorunumBoyu()
        {
            try { return Convert.ToDouble(AcApp.GetSystemVariable("VIEWSIZE"), CultureInfo.InvariantCulture); }
            catch { return -1; }
        }
    }

    // =================================================================
    //  ÇALIŞTIR sonrası denetim: tablodaki her kesit çizildi mi, görünmeyen her eleman bir kesitte mi?
    // =================================================================
    internal static class KesitDenetimi
    {
        private static readonly Regex KesitYazisi = new Regex(@"^SECTION\s+(\S+)$", RegexOptions.IgnoreCase);

        /// <returns>Sorun yoksa null; varsa kısa açıklama (ayrıntı komut satırına).</returns>
        /// <param name="icEleman">İç eleman çizgileri çizilecek: sade modun kesit almadığı elemanlar görünüşte o çizgilerle
        /// durur, eksik sayılmaz.</param>
        internal static string Yap(Form f, bool icEleman = false)
        {
            var parcalar = new List<string>();

            // 1) Tabloda olup çizilmeyen kesit (OutlineDrawing elemanı olmayan kesiti başlığıyla birlikte atlıyor)
            var grid = FormBaglayici.Alan(f, "dataGridView1") as DataGridView;
            var adlar = new List<string>();
            if (grid != null && grid.Columns.Contains("ColumnSection"))
                foreach (DataGridViewRow row in grid.Rows)
                {
                    if (row.IsNewRow) continue;
                    bool dolu = row.Cells.Cast<DataGridViewCell>().Any(c => c.OwningColumn.Name != "ColumnSection" &&
                                                                            c.Value != null && c.Value.ToString().Trim().Length > 0);
                    object v = row.Cells["ColumnSection"].Value;
                    string ad = v == null ? "" : v.ToString().Trim();
                    if (dolu && ad.Length > 0) adlar.Add(ad);
                }
            var cizili = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Document doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc != null && adlar.Count > 0)
            {
                using (doc.LockDocument())
                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    var ms = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(doc.Database), OpenMode.ForRead);
                    foreach (ObjectId id in ms)
                    {
                        if (!Sinif.Yazi(id)) continue;   // nesneyi açmadan sınıfına bak (hız)
                        var t = tr.GetObject(id, OpenMode.ForRead) as DBText;
                        if (t == null || t.TextString == null) continue;
                        Match m = KesitYazisi.Match(t.TextString.Trim());
                        if (m.Success) cizili.Add(m.Groups[1].Value);
                    }
                    tr.Commit();
                }
                List<string> eksik = adlar.Where(a => !cizili.Contains(a)).ToList();
                if (eksik.Count > 0)
                    parcalar.Add(string.Format("Çizilmeyen kesit: {0} (düzleminde hiç eleman yok: kot / nokta değerlerini kontrol edin; " +
                        "mm yazın, binlik ayırıcı kullanmayın)", string.Join(", ", eksik)));
            }

            // 2) Çizilen görünüşlerde görünmeyen ve tablodaki hiçbir kesite düşmeyen eleman (aynası da yoksa)
            string yol = FormBaglayici.YolAl(f);
            KuleVerisi kule = yol.Length > 0 && File.Exists(yol) ? KuleOkuyucu.Oku(f.GetType().Assembly, yol) : null;
            if (kule != null)
            {
                bool on, yan;
                FormBaglayici.GorunusSecimi(f, out on, out yan);
                ICollection<string> gorunen = FormBaglayici.GorunenAnahtarlar(kule, true, true) ?? new List<string>();
                var kesitteki = new HashSet<string>();
                foreach (KesitTanimi t in KesitStili.GridTanimlari(f).Values)
                {
                    List<string> a;
                    KesitSecici.Sec(kule.Uyeler, t, out a);
                    kesitteki.UnionWith(a);
                }
                List<string> kap = AutoSectionDetector.Kapsanmayanlar(kule.Uyeler, gorunen, kesitteki);
                if (icEleman && kap.Count > 0)
                {
                    var ic = new HashSet<string>();
                    if (on) ic.UnionWith(IcElemanPlani.Hesapla(kule.Uyeler, kule.OnYuz, false).Select(c => c.Anahtar));
                    if (yan) ic.UnionWith(IcElemanPlani.Hesapla(kule.Uyeler, kule.YanYuz, true).Select(c => c.Anahtar));
                    var icAile = new HashSet<string>(ic.Where(kule.Uyeler.ContainsKey).Select(k => AutoSectionDetector.AileAnahtari(kule.Uyeler[k])));
                    kap = kap.Where(k => !icAile.Contains(AutoSectionDetector.AileAnahtari(kule.Uyeler[k]))).ToList();
                }
                if (kap.Count > 0)
                {
                    var gruplar = kap.Select(k => kule.Uyeler[k].group_label ?? "?").Distinct().ToList();
                    parcalar.Add(string.Format("{0} eleman ne çizilen görünüşlerde ne de bir kesitte (gruplar: {1}{2}). " +
                        "Kesit sildiyseniz / görünüş seçimini değiştirdiyseniz Oto Kesit'i yeniden çalıştırın.",
                        kap.Count, string.Join(", ", gruplar.Take(8)), gruplar.Count > 8 ? ", ..." : ""));
                    Komutlar.Yaz("\nKesitte olmayan görünmeyen elemanlar: " + string.Join(", ", kap) + "\n");
                }
            }
            if (parcalar.Count == 0) return null;
            string sonuc = string.Join(" | ", parcalar);
            Komutlar.Yaz("\nKESİT DENETİMİ: " + sonuc + "\n");
            return sonuc;
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
        /// <param name="tekKesitPenceresi">Çizimde tek kesit varsa u penceresinin yarı genişliği (kule genişliğinden);
        /// birden çok kesit varsa pencere, çizimdeki kesit başlıkları arasındaki en küçük aralığın yarısıdır.</param>
        /// <param name="pencereV">v (düşey) penceresinin yarı yüksekliği.</param>
        internal static List<KesitCizimi> Topla(Transaction tr, BlockTableRecord ms, string silinecekKatman,
            double tekKesitPenceresi = 15000, double pencereV = 40000)
        {
            var etiketler = new List<KesitCizimi>();
            var cizgiler = new List<KeyValuePair<ObjectId, Point3d[]>>();
            var etiketli = new List<KeyValuePair<string, KeyValuePair<ObjectId, Point3d[]>>>();
            foreach (ObjectId id in ms)
            {
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent == null) continue;
                if (silinecekKatman != null && ent is Dimension && string.Equals(ent.Layer, silinecekKatman, StringComparison.OrdinalIgnoreCase))
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
                if (l != null && !l.Layer.StartsWith("SECTION_MARK", StringComparison.OrdinalIgnoreCase)
                    && !IcElemanlar.BizimKatman(l.Layer))
                {
                    string etiket = KesitDizici.Etiket(l);
                    if (etiket != null) etiketli.Add(new KeyValuePair<string, KeyValuePair<ObjectId, Point3d[]>>(etiket,
                        new KeyValuePair<ObjectId, Point3d[]>(id, new[] { l.StartPoint, l.EndPoint })));
                    else cizgiler.Add(new KeyValuePair<ObjectId, Point3d[]>(id, new[] { l.StartPoint, l.EndPoint }));
                }
            }
            // Kesitler dizildiyse (KesitDizici) çizgiler kesit adıyla işaretli: pencereye bakmadan o kesite
            foreach (var et in etiketli)
            {
                KesitCizimi e = etiketler.FirstOrDefault(x => x.Ad == et.Key);
                if (e == null)
                {
                    // uzatma görünüşünün (aile dosyası) çizgisi hiçbir kesite ait değildir
                    if (!KesitDizici.UzatmaYazisi.IsMatch(et.Key)) cizgiler.Add(et.Value);
                    continue;
                }
                Point3d[] c = et.Value.Value;
                e.Ids.Add(et.Value.Key);
                e.Segs.Add(new[] { c[0].X, e.XZ ? c[0].Z : c[0].Y, c[1].X, e.XZ ? c[1].Z : c[1].Y });
            }

            // u penceresi: OutlineDrawing kesitleri sabit aralıkla (asıl kod 30000, tek DLL kule genişliğine göre) dizer;
            // pencere bu aralığın yarısı (en az 15000). Tek kesitte kule genişliğinden.
            double pencereU = Math.Max(15000, tekKesitPenceresi);
            var uSirali = etiketler.Select(x => x.U).OrderBy(x => x).ToList();
            double enKucukAralik = double.MaxValue;
            for (int i = 1; i < uSirali.Count; i++)
                if (uSirali[i] - uSirali[i - 1] > 1000) enKucukAralik = Math.Min(enKucukAralik, uSirali[i] - uSirali[i - 1]);
            if (enKucukAralik < double.MaxValue) pencereU = Math.Max(15000, enKucukAralik / 2);
            double pv = Math.Max(40000, pencereV);

            foreach (KesitCizimi e in etiketler)
            {
                // Bu etikete ait çizgiler: aynı düzlemde, en yakın etiket bu olan ve pencere içinde kalanlar
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
                    if (Math.Abs(um - e.U) > pencereU || Math.Abs(vm - e.V) > pv) continue;
                    bool baskasi = etiketler.Any(o => o != e && o.XZ == e.XZ && Math.Abs(o.W - e.W) <= 1 &&
                                                      Math.Abs(um - o.U) < Math.Abs(um - e.U) && Math.Abs(vm - o.V) <= pv);
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
            int adet = 0, kesit = 0, gorunus = 0, hatali = 0;

            using (doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var ms = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);
                KatmaniHazirla(db, tr);
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                var olcuKatmani = (LayerTableRecord)tr.GetObject(lt[Katman], OpenMode.ForRead);
                if (olcuKatmani.IsLocked || olcuKatmani.IsOff || olcuKatmani.IsFrozen)
                {
                    olcuKatmani.UpgradeOpen(); olcuKatmani.IsLocked = false; olcuKatmani.IsOff = false; olcuKatmani.IsFrozen = false;
                }
                _metinStili = OlcuMetinStili(db, tr);

                KuleVerisi kule = null;
                try
                {
                    Form f = FormBaglayici.AcikFormuBul();
                    string yol = f != null ? FormBaglayici.YolAl(f) : "";
                    if (yol.Length > 0 && File.Exists(yol)) kule = KuleOkuyucu.Oku(f.GetType().Assembly, yol);
                }
                catch (System.Exception ex) { Komutlar.Yaz("\nÖlçü: .tow okunamadı (" + ex.Message + "); yalnız kesit ölçüleri.\n"); }
                // Görünüşler önce bulunur: başlığı olup yeri bulunamayan görünüş varsa eski ölçüler silinmez
                List<GorunusKonumu> konumlar = kule != null ? GorunusYeri.Bul(tr, ms, kule) : new List<GorunusKonumu>();
                if (kule != null && GorunusYeri.EksikBulundu(konumlar))
                {
                    tr.Commit();
                    return "Ölçü: görünüşün yeri bulunamadı (seçili .tow bu çizimin .tow'u mu?); mevcut ölçülere dokunulmadı.";
                }
                double tekPencere = kule != null ? kule.KesitAraligi / 2.0 : 15000;
                double pencereV = kule != null ? 1.5 * Math.Max(kule.Genislik, kule.Yukseklik) : 40000;

                foreach (KesitCizimi e in KesitCizimleri.Topla(tr, ms, Katman, tekPencere, pencereV))
                {
                    if (e.Segs.Count == 0) continue;
                    List<OlcuTanimi> olculer = OlcuPlani.Hesapla(e.Segs);
                    Matrix3d m = e.XZ
                        ? Matrix3d.Displacement(new Vector3d(0, e.W, 0)) * Matrix3d.Rotation(Math.PI / 2, Vector3d.XAxis, Point3d.Origin)
                        : Matrix3d.Displacement(new Vector3d(0, 0, e.W));
                    foreach (OlcuTanimi o in olculer) { if (OlcuKoy(db, tr, ms, o, m)) adet++; else hatali++; }
                    kesit++;
                }

                // ---- ön (TRANSVERSE) ve yan (LONGITUDINAL) görünüş ölçüleri
                try
                {
                    if (kule != null)
                    {
                        foreach (GorunusKonumu k in konumlar)
                        {
                            List<YuzCizgisi> yuz = k.Yan ? kule.YanYuz : kule.OnYuz;
                            // Ölçüler görünüş başlığının düzleminde: çizildiği hal Y=0 (normal -Y), 2D Aktar sonrası Z=0
                            Matrix3d m = k.IkiD ? Matrix3d.Identity : Matrix3d.Rotation(Math.PI / 2, Vector3d.XAxis, Point3d.Origin);
                            foreach (OlcuTanimi o in GorunusOlcuPlani.Hesapla(yuz, k.Ofs)) { if (OlcuKoy(db, tr, ms, o, m)) adet++; else hatali++; }
                            gorunus++;
                        }
                    }
                }
                catch (System.Exception ex) { Komutlar.Yaz("\nGörünüş ölçüsü konamadı: " + ex.Message + "\n"); }
                tr.Commit();
            }
            string ozet = string.Format("Ölçü: {0} kesit, {1} görünüş", kesit, gorunus);
            if (hatali > 0) ozet += string.Format(" ({0} ölçü AutoCAD tarafından reddedildi: {1})", hatali, _sonOlcuHatasi);
            Komutlar.Yaz(string.Format("\nÖlçü ayrıntı: {0} kesit + {1} görünüşe {2} ölçü (katman {3}).\n", kesit, gorunus, adet, Katman));
            Komutlar.Yaz("\n" + ozet + ".\n");
            return ozet;
        }

        private static string _sonOlcuHatasi;

        /// <returns>false: AutoCAD ölçüyü kabul etmedi (öbür ölçüler etkilenmez)</returns>
        private static bool OlcuKoy(Database db, Transaction tr, BlockTableRecord ms, OlcuTanimi o, Matrix3d m)
        {
            try
            {
                OlcuOlustur(db, tr, ms, o, m);
                return true;
            }
            catch (System.Exception ex)
            {
                _sonOlcuHatasi = ex.Message;
                return false;
            }
        }

        private static void OlcuOlustur(Database db, Transaction tr, BlockTableRecord ms, OlcuTanimi o, Matrix3d m)
        {
            var dim = new RotatedDimension(o.Rotation,
                new Point3d(o.U1, o.V1, 0), new Point3d(o.U2, o.V2, 0), new Point3d(o.UD, o.VD, 0),
                "", db.Dimstyle);
            dim.SetDatabaseDefaults(db);
            if (!_metinStili.IsNull) dim.TextStyleId = _metinStili;   // sabit yükseklikli stil Dimtxt'yi ezmesin
            dim.Layer = Katman;
            dim.Dimtxt = o.Yazi;
            dim.Dimasz = 0.75 * o.Yazi;
            dim.Dimexo = 0.5 * o.Yazi;
            dim.Dimexe = 0.5 * o.Yazi;
            dim.Dimgap = 0.25 * o.Yazi;
            dim.Dimdec = 0;
            dim.Dimtad = 1;
            dim.Dimatfit = 1;   // yer yoksa önce oklar dışarı (yazı ortada kalır); kullanıcının ölçü stilinden bağımsız
            dim.Dimtix = false;
            dim.Dimtih = false;
            dim.Dimtoh = false;
            dim.DimfxlenOn = o.SabitUzatma;   // uzatma çizgisi sabit boy yalnız geometriden uzaktaki köşelerde
            dim.Dimfxlen = o.Uzatma;          // (çizim boyunca uzanan uzun çizgi olmasın); yakındakiler köşeye kadar iner
            dim.Dimscale = 1;     // çizimin ölçü stili ölçekli olsa da yazı boyu korunsun
            dim.Dimlfac = 1;      // ölçülen değer gerçek mm olsun
            dim.Dimtfac = 1;
            dim.Dimlunit = 2;     // ondalık
            dim.Dimrnd = 0;
            if (o.YaziU.HasValue && o.YaziV.HasValue)
            {
                dim.Dimtmove = 2;   // yazı serbest: ok / uzatma çizgileri yerinde kalır
                dim.TextPosition = new Point3d(o.YaziU.Value, o.YaziV.Value, 0);
            }
            try
            {
                dim.TransformBy(m);
                ms.AppendEntity(dim);
            }
            catch
            {
                if (dim.ObjectId.IsNull) dim.Dispose();   // veritabanına girmediyse bellekte kalmasın
                throw;
            }
            tr.AddNewlyCreatedDBObject(dim, true);
        }

        private static ObjectId _metinStili = ObjectId.Null;

        /// <summary>Ölçü stilinin yazı stili sabit yükseklikliyse (ör. şablonda 2,5) AutoCAD Dimtxt'yi yok sayar ve ölçü yazıları
        /// kulede görünmeyecek kadar küçük / dev çıkar. O durumda aynı yazı tipinde, yüksekliği 0 olan OTO_OLCU stili kullanılır.</summary>
        private static ObjectId OlcuMetinStili(Database db, Transaction tr)
        {
            try
            {
                var ds = (DimStyleTableRecord)tr.GetObject(db.Dimstyle, OpenMode.ForRead);
                ObjectId stilId = ds.Dimtxsty.IsNull ? db.Textstyle : ds.Dimtxsty;
                var stil = (TextStyleTableRecord)tr.GetObject(stilId, OpenMode.ForRead);
                if (stil.TextSize <= 1e-9) return ObjectId.Null;
                var tt = (TextStyleTableRecord)null;
                var tst = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
                const string ad = "OTO_OLCU";
                if (tst.Has(ad))
                {
                    tt = (TextStyleTableRecord)tr.GetObject(tst[ad], OpenMode.ForRead);
                    if (tt.TextSize > 1e-9) { tt.UpgradeOpen(); tt.TextSize = 0; }
                    return tt.ObjectId;
                }
                tst.UpgradeOpen();
                tt = new TextStyleTableRecord
                {
                    Name = ad, TextSize = 0, FileName = stil.FileName, BigFontFileName = stil.BigFontFileName,
                    XScale = stil.XScale, ObliquingAngle = stil.ObliquingAngle
                };
                if (string.IsNullOrEmpty(stil.FileName)) tt.Font = stil.Font;
                ObjectId id = tst.Add(tt);
                tr.AddNewlyCreatedDBObject(tt, true);
                return id;
            }
            catch (System.Exception ex) { Komutlar.BirKezYaz("\nÖlçü yazı stili ayarlanamadı: " + ex.Message + "\n"); return ObjectId.Null; }
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
                // Görünüşler bulundukları yerde ve yalnız kendi çizgileriyle aranır: kare gövdede ön ve yan yüz çizgileri
                // aynı ofsette çakışır, birinin grubu öbürüne verilmesin. Bulunamazsa (başlık yok, yüzler özdeş) eski yol.
                var indeksler = new List<YuzEslestirici>();
                List<GorunusKonumu> konumlar = GorunusYeri.Bul(tr, ms, kule);
                foreach (GorunusKonumu k in konumlar)
                    indeksler.Add(new YuzEslestirici(k.Yan ? kule.YanYuz : kule.OnYuz, new[] { k.Ofs }, !k.IkiD, k.IkiD));
                indeksler.Add(new YuzEslestirici(model, new[] { 0 }, true, true));   // YÜKLE'nin 3D modeli ofsetsiz
                if (konumlar.Count == 0)
                    indeksler.Add(new YuzEslestirici(kule.OnYuz.Concat(kule.YanYuz), new[] { 0, 50000, kule.GorunusAraligi }.Distinct().ToArray(), true, true));
                foreach (ObjectId id in ms)
                {
                    var l = tr.GetObject(id, OpenMode.ForRead) as Line;
                    if (l == null || IcElemanlar.BizimKatman(l.Layer)) continue;
                    Point3d a = l.StartPoint, b = l.EndPoint;
                    string grup = null;
                    if (!indeksler.Any(x => x.Bul(a.X, a.Y, a.Z, b.X, b.Y, b.Z, out grup))) continue;
                    try
                    {
                        l.UpgradeOpen();
                        if (Boya(l, grup != null && kule.RedundantGruplar.Contains(grup.Trim()), kesik, duz, olcek)) yuzRed++; else yuzAna++;
                    }
                    catch (System.Exception ex) { hata++; if (ilkHata == null) ilkHata = ex.Message; }
                }

                foreach (KesitCizimi e in KesitCizimleri.Topla(tr, ms, null, kule.KesitAraligi / 2.0, 1.5 * Math.Max(kule.Genislik, kule.Yukseklik)))
                {
                    if (e.Segs.Count == 0) continue;
                    KesitTanimi tanim;
                    if (!tanimlar.TryGetValue(e.Ad, out tanim))
                    {
                        if (!UzatmaCizimleri.SonAdlar.Contains(e.Ad)) tanimsiz.Add(e.Ad);   // aile uzatma kesitleri kendi stiliyle çizildi
                        continue;
                    }
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
        internal static double DesenOlcegi(Database db)
        {
            double lts = db.Ltscale > 1e-9 ? db.Ltscale : 1.0, anno = 1.0;
            try
            {
                if (Convert.ToInt32(AcApp.GetSystemVariable("MSLTSCALE")) != 0 && db.Cannoscale != null && db.Cannoscale.Scale > 1e-9) anno = 1.0 / db.Cannoscale.Scale;
            }
            catch { /* eski sürüm: anotasyon ölçeği yok */ }
            Komutlar.BirKezYaz(string.Format(CultureInfo.InvariantCulture,
                "\nKesik çizgi: LTSCALE {0:0.###}, anotasyon çarpanı {1:0.###} -> nesne çizgi tipi ölçeği {2:0.#####}.\n", lts, anno, 1.0 / (lts * anno)));
            return 1.0 / (lts * anno);
        }

        /// <returns>redundant ise true</returns>
        internal static bool Boya(Entity ent, bool redundant, ObjectId kesik, ObjectId duz, double olcek)
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
        internal static ObjectId KesikCizgiTipi(Database db, Transaction tr)
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
        internal static Dictionary<string, KesitTanimi> GridTanimlari(Form f)
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

    /// <summary>Otomatik kesitler: AutoSectionDetector.DetectHidden + OutlineDrawing'in gerçekte çizeceği elemanlara (KesitSecici)
    /// göre gereksiz kesit ayıklama. (Kesit poligonu %5 büyütüldüğü ve düzlem toleransı 100 mm olduğu için bir kesit hesaplanandan
    /// fazlasını çizer; tek başına hiçbir görünmeyen elemanı göstermeyen kesit tabloya yazılmaz.)</summary>
    public static class OtoKesitBulucu
    {
        /// <param name="sade">Sade (çizimci gibi) kesit. Yalnız ön ve yan görünüş birlikte çizilirken anlamlı: sade modun
        /// kesit almadığı elemanlar görünüşlerde iç eleman çizgisi olarak durur; tek görünüşte tam kapsama gerekir.</param>
        public static List<AutoSectionCandidate> Bul(IDictionary<string, FinalMember> uyeler, ICollection<string> gorunen, out AutoSectionReport rapor,
            bool sade = true)
        {
            List<AutoSectionCandidate> l = AutoSectionDetector.DetectHidden(uyeler, gorunen, new AutoSectionOptions { Sade = sade }, out rapor);
            var secim = new List<ICollection<string>>();
            foreach (AutoSectionCandidate c in l)
            {
                List<string> a;
                KesitSecici.Sec(uyeler, Tanim(c), out a);
                secim.Add(a);
            }
            HashSet<int> at = AutoSectionDetector.GereksizKesitler(uyeler, gorunen, secim);
            if (at.Count > 0)
                Komutlar.Yaz(string.Format("\nOto Kesit: {0} gereksiz kesit atıldı (gösterdiği elemanlar başka kesitlerde de var).\n", at.Count));
            return l.Where((c, i) => !at.Contains(i)).ToList();
        }

        /// <summary>Adayın tabloya yazılan hali (yükseklik "0.###", noktalar FormatPoint ile 3 ondalık).</summary>
        public static KesitTanimi Tanim(AutoSectionCandidate c)
        {
            var t = new KesitTanimi();
            if (c.Kind == AutoSectionKind.Height && c.Height.HasValue) t.Yukseklik = Math.Round(c.Height.Value, 3);
            else foreach (var p in c.Points) t.Noktalar.Add(new[] { Math.Round(p.X, 3), Math.Round(p.Y, 3), Math.Round(p.Z, 3) });
            return t;
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
        public static double PoligonBuyutme = 1.05; // PreparePolygon: kesit çokgeni merkezine göre %5 büyütülür

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

            // OutlineDrawing ile birebir: BuildPlaneFromPoints (P0, P1, Pson), BuildPlaneAxes, PreparePolygon
            double[] p0 = t.Noktalar[0], p1 = t.Noktalar[1], pn = t.Noktalar[t.Noktalar.Count - 1];
            double[] n = Norm(Cross(Sub(p1, p0), Sub(pn, p0)));
            if (n == null) return segs;
            double[] X = { 1, 0, 0 }, Y = { 0, 1, 0 }, Z = { 0, 0, 1 };
            // v: Z'nin düzleme izdüşümü (yatay düzlemde Y'ninki), yukarı; u = v × n, +X yönüne; v = n × u, yukarı
            double[] v = Sub(Z, Kat(n, Dot(Z, n)));
            if (Boy(v) < 1e-9) v = Sub(Y, Kat(n, Dot(Y, n)));
            v = Norm(v);
            if (v == null) return segs;
            if (Dot(v, Z) < 0) v = Kat(v, -1);
            double[] u = Norm(Cross(v, n));
            if (Dot(u, X) < 0) u = Kat(u, -1);
            v = Norm(Cross(n, u));
            if (Dot(v, Z) < 0) v = Kat(v, -1);
            Func<double[], double[]> proj = q => { var d = Sub(q, p0); return new[] { Dot(d, u), Dot(d, v) }; };

            // PreparePolygon: köşeler ortalamalarına göre %5 büyütülür ve açıya göre sıralanır
            var poly = t.Noktalar.Select(proj).ToList();
            double cx = poly.Average(q => q[0]), cy = poly.Average(q => q[1]);
            poly = poly.Select(q => new[] { cx + (q[0] - cx) * PoligonBuyutme, cy + (q[1] - cy) * PoligonBuyutme })
                       .OrderBy(q => Math.Atan2(q[1] - cy, q[0] - cx)).ToList();

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

        /// <summary>OutlineDrawing BuildPlaneAxes'in kopyası: u çizimde sağa (düzlemin yatay doğrultusu, +X'e), v yukarı
        /// (eğim yukarı). Yükseklik kesitinde ya da bozuk düzlemde false.</summary>
        internal static bool Eksenler(KesitTanimi t, out double[] u, out double[] v)
        {
            u = v = null;
            if (t == null || t.Yukseklik.HasValue || t.Noktalar.Count < 3) return false;
            double[] p0 = t.Noktalar[0], p1 = t.Noktalar[1], pn = t.Noktalar[t.Noktalar.Count - 1];
            double[] n = Norm(Cross(Sub(p1, p0), Sub(pn, p0)));
            if (n == null) return false;
            double[] X = { 1, 0, 0 }, Y = { 0, 1, 0 }, Z = { 0, 0, 1 };
            v = Sub(Z, Kat(n, Dot(Z, n)));
            if (Boy(v) < 1e-9) v = Sub(Y, Kat(n, Dot(Y, n)));
            v = Norm(v);
            if (v == null) return false;
            if (Dot(v, Z) < 0) v = Kat(v, -1);
            u = Norm(Cross(v, n));
            if (u == null) return false;
            if (Dot(u, X) < 0) u = Kat(u, -1);
            v = Norm(Cross(n, u));
            if (v == null) return false;
            if (Dot(v, Z) < 0) v = Kat(v, -1);
            return true;
        }

        private static double[] Sub(double[] a, double[] b) { return new[] { a[0] - b[0], a[1] - b[1], a[2] - b[2] }; }
        private static double Dot(double[] a, double[] b) { return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]; }
        private static double[] Cross(double[] a, double[] b) { return new[] { a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2], a[0] * b[1] - a[1] * b[0] }; }
        private static double[] Norm(double[] a) { double l = Math.Sqrt(Dot(a, a)); return l < 1e-9 ? null : new[] { a[0] / l, a[1] / l, a[2] / l }; }
        private static double[] Kat(double[] a, double k) { return new[] { a[0] * k, a[1] * k, a[2] * k }; }
        private static double Boy(double[] a) { return Math.Sqrt(Dot(a, a)); }
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
        /// <summary>Uzatma çizgisi sabit boy mu (false: ölçülen noktaya kadar iner).</summary>
        public bool SabitUzatma = true;
        /// <summary>Yazı sığmıyorsa elle konum (ölçü düzleminde); null: AutoCAD yerleştirir.</summary>
        public double? YaziU, YaziV;
        public double Deger { get { return Math.Abs(Rotation) < 1e-9 ? Math.Abs(U2 - U1) : Math.Abs(V2 - V1); } }
    }

    /// <summary>AutoCAD'den bağımsız ölçü yerleşimi: kesit çizgilerinin dış konturundan
    /// üstte zincir ölçü (kontur köşelerinin u değerleri), sağda toplam derinlik.</summary>
    /// <summary>Kesitlerin ızgara yerleşimi (AutoCAD'den bağımsız): sırayla, kendi boyutlarına göre satırlar halinde. OutlineDrawing
    /// her kesiti en geniş kesit kadar aralıkla tek uzun sıraya diziyordu.</summary>
    public static class KesitDizimi
    {
        public static double EnAzBosluk = 2000;   // kesitler arası en az boşluk (mm)

        /// <param name="kutular">Her kesit (çizgi, etiket, başlık, ölçü dahil) için { u0, v0, u1, v1 }; v yukarı; çizim sırasıyla.</param>
        /// <returns>Her kesit için { du, dv } kaydırma.</returns>
        public static List<double[]> Yerlestir(List<double[]> kutular, out int satir)
        {
            var sonuc = new List<double[]>();
            satir = 0;
            if (kutular == null || kutular.Count == 0) return sonuc;
            double maxW = kutular.Max(k => k[2] - k[0]), maxH = kutular.Max(k => k[3] - k[1]);
            double g = Math.Max(EnAzBosluk, 0.05 * Math.Max(maxW, maxH));
            double alan = kutular.Sum(k => (k[2] - k[0] + g) * (k[3] - k[1] + g));
            double R = Math.Max(maxW, 1.6 * Math.Sqrt(alan));   // satır genişliği: kabaca kare bir blok
            double x0 = kutular.Min(k => k[0]), ust = kutular.Max(k => k[3]);
            double x = x0, satirUst = ust, satirH = 0;
            satir = 1;
            foreach (double[] k in kutular)
            {
                double w = k[2] - k[0], h = k[3] - k[1];
                if (x > x0 && x + w > x0 + R) { satirUst -= satirH + g; x = x0; satirH = 0; satir++; }
                sonuc.Add(new[] { x - k[0], satirUst - k[3] });
                x += w + g;
                satirH = Math.Max(satirH, h);
            }
            return sonuc;
        }
    }

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
                var zincir = new List<OlcuTanimi>();
                for (int i = 0; i + 1 < us.Count; i++)
                    zincir.Add(new OlcuTanimi
                    {
                        U1 = us[i][0], V1 = us[i][1], U2 = us[i + 1][0], V2 = us[i + 1][1],
                        UD = (us[i][0] + us[i + 1][0]) / 2, VD = vd, Rotation = 0, Yazi = yazi, Uzatma = uzatma,
                        // köşe ölçü çizgisine yakınsa uzatma çizgisi köşeye kadar iner; derindeki köşede sabit boy
                        SabitUzatma = Math.Max(vd - us[i][1], vd - us[i + 1][1]) > 3 * aralik
                    });
                YazilariYerlestir(zincir, vd);
                sonuc.AddRange(zincir);
            }

            // Sağ: toplam derinlik (en üst ve en alt köşelerden sağdakiler)
            if (maxV - minV >= EnKisa)
            {
                var ust = hull.Where(p => p[1] >= maxV - 1e-6).OrderByDescending(p => p[0]).First();
                var alt = hull.Where(p => p[1] <= minV + 1e-6).OrderByDescending(p => p[0]).First();
                sonuc.Add(new OlcuTanimi
                {
                    U1 = ust[0], V1 = ust[1], U2 = alt[0], V2 = alt[1],
                    UD = maxU + aralik, VD = (maxV + minV) / 2, Rotation = Math.PI / 2, Yazi = yazi, Uzatma = uzatma,
                    SabitUzatma = Math.Max(maxU + aralik - ust[0], maxU + aralik - alt[0]) > 3 * aralik
                });
            }
            return sonuc;
        }

        /// <summary>Yazı iki uzatma çizgisi arasına sığar mı (yazı + iki yandaki boşluk + pay). Oklar sığmazsa AutoCAD
        /// okları dışarı alır, yazı ortada kalır (DIMATFIT = 1); yazı da sığmıyorsa yeri elle verilir.</summary>
        public static bool YaziSigar(OlcuTanimi o)
        {
            return o.Deger >= YaziGenisligi(o) + 0.7 * o.Yazi;
        }

        /// <summary>Yazı genişliği tahmini (mm): rakam başına ~0,75 yazı yüksekliği.</summary>
        public static double YaziGenisligi(OlcuTanimi o)
        {
            int n = Math.Round(o.Deger).ToString(CultureInfo.InvariantCulture).Length;
            return n * 0.75 * o.Yazi;
        }

        /// <summary>Zincirde yazısı iki ok arasına sığmayan kısa parçaların yazısı komşusuyla çakışmasın: ilk parça
        /// sola, son parça sağa, aradakiler ölçü çizgisinin üstüne (art arda kısalar sırayla bir kat daha yukarı).</summary>
        public static void YazilariYerlestir(List<OlcuTanimi> zincir, double vd)
        {
            int kat = 0;
            for (int i = 0; i < zincir.Count; i++)
            {
                OlcuTanimi o = zincir[i];
                double w = YaziGenisligi(o);
                if (YaziSigar(o)) { kat = 0; continue; }
                double sol = Math.Min(o.U1, o.U2), sag = Math.Max(o.U1, o.U2);
                if (i == 0 && zincir.Count > 1) { o.YaziU = sol - 0.75 * o.Yazi - w / 2; o.YaziV = vd + 0.5 * o.Yazi; kat = 0; }
                else if (i == zincir.Count - 1 && zincir.Count > 1) { o.YaziU = sag + 0.75 * o.Yazi + w / 2; o.YaziV = vd + 0.5 * o.Yazi; }
                else if (zincir.Count > 1) { kat = kat % 2 + 1; o.YaziU = (sol + sag) / 2; o.YaziV = vd + 0.5 * o.Yazi + kat * 1.5 * o.Yazi; }
            }
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

    // =================================================================
    //  Görünüşlerin çizimdeki yeri
    // =================================================================
    internal static class GorunusYeri
    {
        /// <summary>Son Bul çağrısında model alanında bulunan görünüş başlığı ("TRANSVERSE / LONGITUDINAL FACE") sayısı.</summary>
        internal static int SonBaslikSayisi;

        /// <summary>Başlığı olduğu hâlde yeri bulunamayan görünüş var mı (ör. .tow değişti, çizim elle değişti)? Öyleyse
        /// o görünüşün eski iç eleman çizgileri / ölçüleri silinmemeli: yerine yenisi konamaz.</summary>
        internal static bool EksikBulundu(List<GorunusKonumu> konumlar)
        {
            return konumlar.Count == 0 ? SonBaslikSayisi > 0 : konumlar.Count < SonBaslikSayisi;
        }

        internal const string OnBaslik = "TRANSVERSE FACE", YanBaslik = "LONGITUDINAL FACE";

        /// <summary>Model alanındaki çizgi ve görünüş başlıklarından ön/yan görünüşün yerini bulur (bkz. GorunusKonumBulucu).
        /// İç eleman katmanındaki çizgiler hesaba katılmaz.</summary>
        private static List<GorunusKonumu> _onbellek;
        private static bool _onbellekAcik;

        /// <summary>ÇALIŞTIR sonrası adımlar boyunca görünüş yerleri bir kez bulunur (adımlar yalnız kendi katmanlarına
        /// çizgi / ölçü ekler ve çizgi tipi değiştirir; görünüş çizgileri değişmez).</summary>
        internal static void OnbellekAc() { _onbellekAcik = true; _onbellek = null; }
        internal static void OnbellekKapat() { _onbellekAcik = false; _onbellek = null; }

        internal static List<GorunusKonumu> Bul(Transaction tr, BlockTableRecord ms, KuleVerisi kule)
        {
            if (_onbellekAcik && _onbellek != null) return new List<GorunusKonumu>(_onbellek);
            List<GorunusKonumu> l = BulTara(tr, ms, kule);
            if (_onbellekAcik) _onbellek = new List<GorunusKonumu>(l);
            return l;
        }

        private static List<GorunusKonumu> BulTara(Transaction tr, BlockTableRecord ms, KuleVerisi kule)
        {
            var cizgiler = new List<double[]>();
            var basliklar = new List<GorunusBasligi>();
            foreach (ObjectId id in ms)
            {
                if (id.IsErased) continue;
                if (!Sinif.Cizgi(id) && !Sinif.Yazi(id)) continue;   // ölçü, blok vb. açılmaz (hız)
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent == null) continue;
                var l = ent as Line;
                if (l != null)
                {
                    if (IcElemanlar.BizimKatman(l.Layer)) continue;
                    cizgiler.Add(new[] { l.StartPoint.X, l.StartPoint.Y, l.StartPoint.Z, l.EndPoint.X, l.EndPoint.Y, l.EndPoint.Z });
                    continue;
                }
                var t = ent as DBText;
                if (t == null || t.TextString == null) continue;
                string s = t.TextString.Trim();
                bool on = s.Equals(OnBaslik, StringComparison.OrdinalIgnoreCase), yan = s.Equals(YanBaslik, StringComparison.OrdinalIgnoreCase);
                if (!on && !yan) continue;
                Point3d p = t.HorizontalMode == TextHorizontalMode.TextLeft && t.VerticalMode == TextVerticalMode.TextBase
                    ? t.Position : t.AlignmentPoint;
                Vector3d n = t.Normal;
                if (Math.Abs(n.Y) > 0.9) basliklar.Add(new GorunusBasligi { Yan = yan, X = p.X, IkiD = false });
                else if (Math.Abs(n.Z) > 0.9) basliklar.Add(new GorunusBasligi { Yan = yan, X = p.X, IkiD = true });
            }
            SonBaslikSayisi = basliklar.Count;
            List<GorunusKonumu> sonuc = GorunusKonumBulucu.Bul(kule.OnYuz, kule.YanYuz, basliklar, cizgiler, kule.GorunusAraligi);
            foreach (GorunusKonumu k in sonuc)
                Komutlar.BirKezYaz(string.Format("\nGörünüş: {0} ofset {1}{2}, {3}/{4} çizgi ({5}).\n", k.Yan ? YanBaslik : OnBaslik,
                    k.Ofs, k.IkiD ? " (2D)" : "", k.Eslesen, k.Toplam, k.Kaynak));
            return sonuc;
        }
    }

    /// <summary>Bir görünüşün çizimdeki yeri.</summary>
    public class GorunusKonumu
    {
        public bool Yan;        // false: TRANSVERSE FACE (ön), true: LONGITUDINAL FACE (yan)
        public int Ofs;         // OutlineDrawing ofseti: çizim x = görünüş x + Ofs
        public bool IkiD;       // 2D Aktar sonrası hal: (x, z, 0)
        public int Eslesen, Toplam;
        public string Kaynak;
    }

    /// <summary>Görünüş başlığı yazısı: OutlineDrawing görünüşün x ortasına (normal -Y) yazar.</summary>
    public class GorunusBasligi
    {
        public bool Yan;
        public double X;
        public bool IkiD;
    }

    /// <summary>
    /// Ön ve yan görünüşün çizimdeki yerini bulur. OutlineDrawing ön görünüşü ofset 0'a, yan görünüşü ön de çizildiyse
    /// 50000'e, çizilmediyse 0'a koyar ve başlığını görünüşün x ortasına yazar. Kare gövdeli kulelerde yan yüz çizgilerinin
    /// çoğu ön yüz çizgileriyle aynı ofsette birebir çakışır (YAA-R1 195'te 136, Tangent Tower 198'de 186). Yalnız çizgi
    /// eşleştirmesine bakmak, sadece ön görünüş çizildiğinde onu yan görünüş de sanıyor; yan görünüşün ölçüleri ve iç
    /// elemanları ön görünüşün üstüne çiziliyordu. Bu yüzden:
    ///  1. Başlık varsa yer başlıktan hesaplanır (ofset = başlık x - görünüşün x ortası, hal başlığın normalinden) ve
    ///     görünüş çizgilerinin en az yarısının orada olduğu doğrulanır.
    ///  2. Başlık yoksa yalnız o yüze özgü (öbür yüzle çakışmayan) çizgilere bakılır.
    ///  3. Özgü çizgi yoksa (yüzler özdeş) görünüş belirsiz sayılır ve kullanılmaz.
    /// İki görünüş aynı yere düşerse başlıktan bulunan (yoksa daha çok eşleşen) kalır.
    /// </summary>
    public static class GorunusKonumBulucu
    {
        public static double EsikOrani = 0.5;
        public static int EnAzOzgu = 3;
        public static List<GorunusKonumu> Bul(List<YuzCizgisi> on, List<YuzCizgisi> yan, List<GorunusBasligi> basliklar, List<double[]> cizgiler,
            int gorunusAraligi = 50000)
        {
            int[] ofsetler = new[] { 0, 50000, gorunusAraligi }.Distinct().ToArray();
            var cizim = new YuzEslestirici(cizgiler.Select(c => new YuzCizgisi { P = c }), new[] { 0 }, true, false);
            var adaylar = new List<GorunusKonumu>();
            foreach (bool y in new[] { false, true })
            {
                List<YuzCizgisi> F = y ? yan : on, G = y ? on : yan;
                if (F == null || F.Count == 0) continue;
                GorunusKonumu k = BasliktanBul(F, y, basliklar, cizim) ?? OzguCizgilerdenBul(F, G, y, cizim, ofsetler);
                if (k != null) adaylar.Add(k);
            }
            var sonuc = new List<GorunusKonumu>();
            foreach (GorunusKonumu k in adaylar.OrderBy(k => k.Kaynak == KaynakBaslik ? 0 : 1).ThenByDescending(k => (double)k.Eslesen / Math.Max(1, k.Toplam)))
                if (!sonuc.Any(o => o.Ofs == k.Ofs && o.IkiD == k.IkiD)) sonuc.Add(k);
            return sonuc.OrderBy(k => k.Yan).ToList();
        }

        public const string KaynakBaslik = "başlıktan", KaynakOzgu = "yüze özgü çizgilerden";

        private static GorunusKonumu BasliktanBul(List<YuzCizgisi> F, bool yan, List<GorunusBasligi> basliklar, YuzEslestirici cizim)
        {
            if (basliklar == null) return null;
            double cx = (F.Min(c => Math.Min(c.P[0], c.P[3])) + F.Max(c => Math.Max(c.P[0], c.P[3]))) / 2;
            GorunusKonumu enIyi = null;
            foreach (GorunusBasligi b in basliklar.Where(b => b.Yan == yan))
            {
                int ofs = (int)Math.Round(b.X - cx);
                int n = Say(F, ofs, b.IkiD, cizim);
                if (n < EsikOrani * F.Count || (enIyi != null && n <= enIyi.Eslesen)) continue;
                enIyi = new GorunusKonumu { Yan = yan, Ofs = ofs, IkiD = b.IkiD, Eslesen = n, Toplam = F.Count, Kaynak = KaynakBaslik };
            }
            return enIyi;
        }

        private static GorunusKonumu OzguCizgilerdenBul(List<YuzCizgisi> F, List<YuzCizgisi> G, bool yan, YuzEslestirici cizim, int[] ofsetler)
        {
            // Özgü çizgiler ofsetten ve halden bağımsız (iki yüz aynı dönüşümle kayar): ofset 0, çizildiği hal
            List<YuzCizgisi> ozgu = F;
            if (G != null && G.Count > 0)
            {
                var g = new YuzEslestirici(G, new[] { 0 }, true, false);
                string grup;
                ozgu = F.Where(c => !g.Bul(c.P[0], c.P[1], c.P[2], c.P[3], c.P[4], c.P[5], out grup)).ToList();
            }
            if (ozgu.Count < EnAzOzgu) return null;
            GorunusKonumu enIyi = null;
            foreach (int ofs in ofsetler)
                foreach (bool iki in new[] { false, true })
                {
                    int n = Say(ozgu, ofs, iki, cizim);
                    if (n < EsikOrani * ozgu.Count || (enIyi != null && n <= enIyi.Eslesen)) continue;
                    enIyi = new GorunusKonumu { Yan = yan, Ofs = ofs, IkiD = iki, Eslesen = n, Toplam = ozgu.Count, Kaynak = KaynakOzgu };
                }
            return enIyi;
        }

        /// <summary>Görünüş çizgilerinden kaçı çizimde (ofs, hal) yerinde var.</summary>
        public static int Say(List<YuzCizgisi> F, int ofs, bool ikiD, YuzEslestirici cizim)
        {
            int n = 0;
            string g;
            foreach (YuzCizgisi c in F)
            {
                double[] p = c.P;
                double[] q = ikiD ? YuzEslestirici.IkiD(p[0] + ofs, p[1], p[2], p[3] + ofs, p[4], p[5])
                                  : new[] { p[0] + ofs, p[1], p[2], p[3] + ofs, p[4], p[5] };
                bool var = cizim.Bul(q[0], q[1], q[2], q[3], q[4], q[5], out g);
                if (var) n++;
            }
            return n;
        }
    }

    // =================================================================
    //  Görünüşlere iç elemanların izdüşümü (PLS-TOWER görünüşündeki gibi)
    // =================================================================
    internal static class IcElemanlar
    {
        internal const string Katman = "OTO_IC_ELEMAN";
        /// <summary>Uçtan görülen (bakış yönüne yakın / gövde dışındaki) elemanların izdüşümü: ör. yan görünüşte travers.
        /// Görünüşte eksik çizgi olmasın diye çizilir; her ÇALIŞTIR'da görünür yapılır.</summary>
        internal const string KatmanUc = "OTO_IC_ELEMAN_UC";

        internal static bool BizimKatman(string katman)
        {
            return string.Equals(katman, Katman, StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(katman, KatmanUc, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>Ön ve yan görünüşe, yüzde olmayan elemanların görünüşte yeni çizgi oluşturan izdüşümlerini ekler.
        /// Önceki çalıştırmanın çizgileri silinir.</summary>
        /// <summary>Son Ciz'de kullanıcının görmesi gereken bir durum (görünüş bulunamadı, çizgi eklenemedi); yoksa null.</summary>
        internal static string SonUyari;

        internal static string Ciz(Form f, bool stil)
        {
            SonUyari = null;
            Document doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return "İç eleman: açık çizim yok.";
            Database db = doc.Database;
            string yol = FormBaglayici.YolAl(f);
            if (yol.Length == 0 || !File.Exists(yol)) return "İç eleman: .tow dosyası seçili değil.";
            KuleVerisi kule = KuleOkuyucu.Oku(f.GetType().Assembly, yol);
            if (kule == null) return "İç eleman: .tow okunamadı.";
            int adet = 0, adetUc = 0, gorunus = 0, plan = 0, hatali = 0, adetOn = 0, adetYan = 0;
            string ilkHata = null;
            var acilan = new List<string>();

            using (doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var ms = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                foreach (string ad in new[] { Katman, KatmanUc })
                {
                    if (!lt.Has(ad))
                    {
                        if (!lt.IsWriteEnabled) lt.UpgradeOpen();
                        var ltr = new LayerTableRecord { Name = ad };
                        lt.Add(ltr);
                        tr.AddNewlyCreatedDBObject(ltr, true);
                    }
                    var katman = (LayerTableRecord)tr.GetObject(lt[ad], OpenMode.ForRead);
                    // OutlineDrawing kendi katmanlarını her çizimde açar (EnsureLayer); kapalı / dondurulmuş katmanda çizgiler
                    // çizilir ama görünmez, görünüş eksik kalırdı. İç eleman kutusu işaretliyse katman görünür yapılır.
                    if (katman.IsLocked || katman.IsOff || katman.IsFrozen)
                    {
                        if (katman.IsOff || katman.IsFrozen) acilan.Add(ad);
                        katman.UpgradeOpen(); katman.IsLocked = false; katman.IsOff = false; katman.IsFrozen = false;
                    }
                }

                // Görünüşler önce bulunur: bulunamazsa eski çizgiler silinmez (yerine yenisi konamaz)
                List<GorunusKonumu> konumlar = GorunusYeri.Bul(tr, ms, kule);
                if (konumlar.Count == 0 || GorunusYeri.EksikBulundu(konumlar))
                {
                    tr.Commit();
                    if (konumlar.Count == 0 && GorunusYeri.SonBaslikSayisi == 0) return "İç eleman: çizimde görünüş yok.";
                    SonUyari = "İç eleman çizgileri eklenmedi: görünüşün yeri bulunamadı (seçili .tow bu çizimin .tow'u mu?). " +
                               "Görünüşte bazı elemanlar eksik görünebilir.";
                    return "İç eleman: görünüşün yeri bulunamadı (seçili .tow bu çizimin .tow'u mu?); mevcut iç eleman çizgilerine dokunulmadı.";
                }

                // Yalnız çizgiler silinir: geçerli katman (CLAYER) bizim katmanımızken OutlineDrawing'in koyduğu
                // "SECTION x" / "... FACE" başlıkları da bu katmana düşer, onlar silinmesin.
                foreach (ObjectId id in ms)
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead) as Line;
                    if (ent == null || !BizimKatman(ent.Layer)) continue;
                    try { ent.UpgradeOpen(); ent.Erase(); } catch { }
                }

                ObjectId kesik = stil ? KesitStili.KesikCizgiTipi(db, tr) : ObjectId.Null;
                ObjectId duz = SymbolUtilityServices.GetLinetypeContinuousId(db);
                double olcek = stil ? KesitStili.DesenOlcegi(db) : 1;
                foreach (GorunusKonumu k in konumlar)
                {
                    List<YuzCizgisi> yuz = k.Yan ? kule.YanYuz : kule.OnYuz;
                    gorunus++;
                    foreach (YuzCizgisi c in IcElemanPlani.Hesapla(kule.Uyeler, yuz, k.Yan))
                    {
                        plan++;
                        // Bir çizgideki hata (ör. çizgi tipi) öbürlerini düşürmesin: her çizgi ayrı korunur
                        Line line = null;
                        try
                        {
                            line = new Line(Nokta(c.P, 0, k.Ofs, k.IkiD), Nokta(c.P, 3, k.Ofs, k.IkiD));
                            line.SetDatabaseDefaults(db);
                            line.Layer = c.Uc ? KatmanUc : Katman;
                            if (stil)
                            {
                                try { KesitStili.Boya(line, c.Grup != null && kule.RedundantGruplar.Contains(c.Grup.Trim()), kesik, duz, olcek); }
                                catch (System.Exception ex) { if (ilkHata == null) ilkHata = ex.Message; }
                            }
                            ms.AppendEntity(line);
                            tr.AddNewlyCreatedDBObject(line, true);
                            adet++;
                            if (c.Uc) adetUc++;
                            if (k.Yan) adetYan++; else adetOn++;
                        }
                        catch (System.Exception ex)
                        {
                            hatali++;
                            if (ilkHata == null) ilkHata = ex.Message;
                            if (line != null && line.ObjectId.IsNull) line.Dispose();
                        }
                    }
                }
                tr.Commit();
            }
            Komutlar.Yaz(string.Format("\nİç eleman: {0} görünüşe {1} izdüşüm çizgisi (katman {2}; {3} tanesi uçtan görülen, katman {4}).\n",
                gorunus, adet, Katman, adetUc, KatmanUc));
            if (hatali > 0)
                SonUyari = string.Format("İç eleman: {0} / {1} çizgi eklenemedi ({2}).", hatali, plan, ilkHata);
            else if (ilkHata != null)
                Komutlar.Yaz("\nİç eleman çizgi tipi uyarısı: " + ilkHata + "\n");
            if (acilan.Count > 0) Komutlar.Yaz("\nİç eleman: kapalı / dondurulmuş katman açıldı: " + string.Join(", ", acilan) + ".\n");
            return string.Format("İç eleman: {0} çizgi (ön {1}, yan {2}){3}", adet, adetOn, adetYan,
                acilan.Count > 0 ? " - kapalı katman açıldı: " + string.Join(", ", acilan) : "");
        }

        // OutlineDrawing görünüşü Point3d(x + ofset, y, z) ile çizer; 2D Aktar sonrası (x, z, 0)
        private static Point3d Nokta(double[] p, int i, int ofs, bool ikiD)
        {
            double u = p[i] + ofs, d = p[i + 1], z = p[i + 2];
            return ikiD ? new Point3d(u, z, 0) : new Point3d(u, d, z);
        }
    }

    /// <summary>Ön görünüşte bir kesit işareti: çizgiler ve harf (görünüş koordinatı u = x, v = z; ofset eklenmemiş).</summary>
    public class KesitIsareti
    {
        public string Ad;
        public List<double[]> Cizgiler = new List<double[]>();   // { u1, v1, u2, v2 }
        public double YaziU, YaziV, YaziBoyu;
        public double[] Kutu;                                     // { u0, v0, u1, v1 }
    }

    /// <summary>
    /// Çizimci gibi kesit işaretleri. OutlineDrawing her kesit için ön görünüşe L oklu bloklar koyar; eğik / parça
    /// kesitlerde bunlar gövdenin içine, eleman yazılarının üstüne ve birbirinin üstüne düşüyor. Burada işaretler kulenin
    /// dışına konur:
    ///  - yatay kesit: o kotta kule silüetinin iki yanında kısa kesit çizgisi, uçta aşağı bakan ok ve harf;
    ///  - ön görünüşte çizgi olarak görünen düzlem (travers üst düzlemi, yüz şeridi, yatay parça): düzlem izinin kuleden
    ///    uzak ucunda, iz doğrultusunda dışa uzanan çizgi, bakış yönünde ok;
    ///  - ön görünüşte alan olarak görünen düzlem: zemine inen bacak iç düzleminde kesit hattı üst-iç köşeden ayağa (bacak
    ///    iç çizgisi boyunca), iki uçta işaret, ok içeriden dışa; gövdedeki köşegen düzlemde ağırlık merkezi kotunda,
    ///    kesitin bulunduğu yanda silüetin dışından kuleyi gösteren ok.
    /// Birbirine binen işaret dışa kaydırılır.
    /// </summary>
    public static class KesitIsaretPlani
    {
        public static List<KesitIsareti> Hesapla(IList<YuzCizgisi> onYuz, IDictionary<string, KesitTanimi> tanimlar)
        {
            var sonuc = new List<KesitIsareti>();
            if (onYuz == null || onYuz.Count == 0 || tanimlar == null || tanimlar.Count == 0) return sonuc;
            double zMin = onYuz.Min(c => Math.Min(c.P[2], c.P[5])), zMax = onYuz.Max(c => Math.Max(c.P[2], c.P[5]));
            double xMin = onYuz.Min(c => Math.Min(c.P[0], c.P[3])), xMax = onYuz.Max(c => Math.Max(c.P[0], c.P[3]));
            double th = Math.Max(150, Math.Min(500, Math.Round((zMax - zMin) / 160 / 50) * 50));
            double bosluk = th, kol = 3 * th, ok = 1.5 * th;
            var dolu = new List<double[]>();

            foreach (var kv in tanimlar.OrderByDescending(k => Kot(k.Value)))
            {
                string ad = kv.Key; KesitTanimi t = kv.Value;
                if (t.Yukseklik.HasValue)
                {
                    double h = t.Yukseklik.Value, a, b;
                    Kenar(onYuz, h, xMin, xMax, out a, out b);
                    Ekle(sonuc, dolu, ad, a - bosluk, h, -1, 0, 0, -1, kol, ok, th);
                    Ekle(sonuc, dolu, ad, b + bosluk, h, 1, 0, 0, -1, kol, ok, th);
                    continue;
                }
                if (t.Noktalar.Count < 3) continue;
                double[] p0 = t.Noktalar[0], p1 = t.Noktalar[1], p2 = t.Noktalar[2];
                double ax = p1[0] - p0[0], ay = p1[1] - p0[1], az = p1[2] - p0[2], bx = p2[0] - p0[0], by = p2[1] - p0[1], bz = p2[2] - p0[2];
                double nx = ay * bz - az * by, ny = az * bx - ax * bz, nz = ax * by - ay * bx, nl = Math.Sqrt(nx * nx + ny * ny + nz * nz);
                if (nl < 1e-9) continue;
                nx /= nl; ny /= nl; nz /= nl;
                if (nz < 0) { nx = -nx; ny = -ny; nz = -nz; }
                double cx = t.Noktalar.Average(p => p[0]), cz = t.Noktalar.Average(p => p[2]);
                if (Math.Abs(ny) < 0.2)
                {
                    // izi ön görünüşte çizgi: doğrultu t = (nz, -nx) (XZ'de), uçlar dörtgenin bu doğrultudaki uç izdüşümleri
                    double hl = Math.Sqrt(nx * nx + nz * nz), tu = nz / hl, tv = -nx / hl;
                    double s0 = t.Noktalar.Min(p => (p[0] - cx) * tu + (p[2] - cz) * tv), s1 = t.Noktalar.Max(p => (p[0] - cx) * tu + (p[2] - cz) * tv);
                    double u0 = cx + tu * s0, v0 = cz + tv * s0, u1 = cx + tu * s1, v1 = cz + tv * s1;
                    // kuleden uzak uç (eşitse üstteki)
                    bool birinci = Math.Abs(Math.Abs(u0) - Math.Abs(u1)) > th ? Math.Abs(u0) > Math.Abs(u1) : v0 > v1;
                    double eu = birinci ? u0 : u1, ev = birinci ? v0 : v1, du = birinci ? -tu : tu, dv = birinci ? -tv : tv;
                    // bakış: eğik düzlemde üstten (-n), düşeyde kule eksenine doğru
                    double bu, bv;
                    if (nz > 0.1) { bu = -nx / hl; bv = -nz / hl; }
                    else { bu = eu > 0 ? -1 : 1; bv = 0; }
                    Ekle(sonuc, dolu, ad, eu + du * bosluk, ev + dv * bosluk, du, dv, bu, bv, kol, ok, th);
                }
                else
                {
                    // ön görünüşte alan olarak görünen düzlem (ör. bacağın iç / kalça düzlemi): çizimci kesit hattını düzlemin
                    // üst-iç köşesinden alt-dış köşesine (bacak iç çizgisi boyunca) çizer; işaret iki uçta, ok içeriden
                    // dışa (kesite içeriden bakılır)
                    var kose = t.Noktalar.OrderByDescending(p => p[2]).ToList();
                    double[] ust = kose.Take(2).OrderBy(p => Math.Abs(p[0])).First(), alt = kose.Skip(2).OrderByDescending(p => Math.Abs(p[0])).First();
                    double lu = alt[0] - ust[0], lv = alt[2] - ust[2], ll = Math.Sqrt(lu * lu + lv * lv);
                    // yalnız zemine inen (bacak) düzlemde; gövdedeki köşegen düzlemlerin üst-iç köşesi gövdenin ortasına düşer
                    bool bacak = alt[2] <= zMin + 2 * th;
                    if (bacak)
                    {
                        // dörtgen köşesi elemanların dışında kalabilir: alt uç bacağın ayağı (silüetin en alttaki ucu)
                        double fa, fb;
                        Kenar(onYuz, zMin, xMin, xMax, out fa, out fb);
                        alt = new[] { cx >= 0 ? fb : fa, 0, zMin };
                        lu = alt[0] - ust[0]; lv = alt[2] - ust[2]; ll = Math.Sqrt(lu * lu + lv * lv);
                    }
                    if (ll < 1e-6 || !bacak)
                    {
                        double a, b;
                        Kenar(onYuz, cz, xMin, xMax, out a, out b);
                        bool sag = cx >= 0;
                        Ekle(sonuc, dolu, ad, sag ? b + bosluk : a - bosluk, cz, sag ? 1 : -1, 0, sag ? -1 : 1, 0, kol, ok, th);
                        continue;
                    }
                    lu /= ll; lv /= ll;
                    double pu = lv, pv = -lu;
                    if (pu * (cx >= 0 ? 1 : -1) < 0) { pu = -pu; pv = -pv; }
                    Ekle(sonuc, dolu, ad, ust[0] - lu * bosluk, ust[2] - lv * bosluk, -lu, -lv, pu, pv, 1.5 * th, ok, th);
                    Ekle(sonuc, dolu, ad, alt[0] + lu * bosluk, alt[2] + lv * bosluk, lu, lv, pu, pv, 1.5 * th, ok, th);
                }
            }
            return sonuc;
        }

        private static double Kot(KesitTanimi t)
        {
            return t.Yukseklik ?? (t.Noktalar.Count > 0 ? t.Noktalar.Average(p => p[2]) : 0);
        }

        /// <summary>Ön yüz silüetinin z kotundaki sol / sağ sınırı (o kotu kesen çizgilerden).</summary>
        private static void Kenar(IList<YuzCizgisi> onYuz, double z, double xMin, double xMax, out double a, out double b)
        {
            a = double.MaxValue; b = double.MinValue;
            foreach (var c in onYuz)
            {
                double z1 = c.P[2], z2 = c.P[5];
                if (z < Math.Min(z1, z2) - 1 || z > Math.Max(z1, z2) + 1) continue;
                if (Math.Abs(z2 - z1) < 1) { a = Math.Min(a, Math.Min(c.P[0], c.P[3])); b = Math.Max(b, Math.Max(c.P[0], c.P[3])); continue; }
                double x = c.P[0] + (c.P[3] - c.P[0]) * (z - z1) / (z2 - z1);
                a = Math.Min(a, x); b = Math.Max(b, x);
            }
            if (a > b) { a = xMin; b = xMax; }
        }

        /// <summary>(u, v) başlangıcından d yönünde kol boyu kesit çizgisi; ucunda b yönünde ok ve harf. Dolu bir yere
        /// binerse d yönünde dışa kaydırılır.</summary>
        private static void Ekle(List<KesitIsareti> sonuc, List<double[]> dolu, string ad, double u, double v,
            double du, double dv, double bu, double bv, double kol, double ok, double th)
        {
            double dl = Math.Sqrt(du * du + dv * dv), bl = Math.Sqrt(bu * bu + bv * bv);
            if (dl < 1e-9) { du = 1; dv = 0; dl = 1; }
            if (bl < 1e-9) { bu = 0; bv = -1; bl = 1; }
            du /= dl; dv /= dl; bu /= bl; bv /= bl;
            KesitIsareti s = null;
            for (int deneme = 0; deneme < 12; deneme++)
            {
                double su = u + du * deneme * 1.5 * th, sv = v + dv * deneme * 1.5 * th;
                double eu = su + du * kol, ev = sv + dv * kol;          // kolun dış ucu
                double ou = eu + bu * ok, ov = ev + bv * ok;            // ok ucu
                double qu = -bv, qv = bu;                               // ok başının yanları
                s = new KesitIsareti { Ad = ad, YaziBoyu = th };
                s.Cizgiler.Add(new[] { su, sv, eu, ev });
                s.Cizgiler.Add(new[] { eu, ev, ou, ov });
                s.Cizgiler.Add(new[] { ou, ov, ou - bu * 0.6 * th + qu * 0.3 * th, ov - bv * 0.6 * th + qv * 0.3 * th });
                s.Cizgiler.Add(new[] { ou, ov, ou - bu * 0.6 * th - qu * 0.3 * th, ov - bv * 0.6 * th - qv * 0.3 * th });
                // harf: kolun dış ucunun ötesinde, oka ters tarafta
                s.YaziU = eu + du * 0.8 * th - bu * 0.6 * th;
                s.YaziV = ev + dv * 0.8 * th - bv * 0.6 * th - 0.5 * th;
                double yg = Math.Max(1, ad.Length) * 0.8 * th;
                var k = new[] { s.YaziU - yg / 2, s.YaziV, s.YaziU + yg / 2, s.YaziV + th };
                foreach (var c in s.Cizgiler) k = new[] { Math.Min(k[0], Math.Min(c[0], c[2])), Math.Min(k[1], Math.Min(c[1], c[3])),
                                                           Math.Max(k[2], Math.Max(c[0], c[2])), Math.Max(k[3], Math.Max(c[1], c[3])) };
                s.Kutu = k;
                double p = 0.3 * th;
                if (!dolu.Any(o => k[0] < o[2] + p && o[0] < k[2] + p && k[1] < o[3] + p && o[1] < k[3] + p)) break;
            }
            dolu.Add(s.Kutu);
            sonuc.Add(s);
        }
    }

    /// <summary>Aile dosyasında UzatmaPlani'ndaki çizimleri (uzatma görünüşleri, K / L kesitleri) model alanına, mevcut
    /// çizimin sağına çizer. Çizgiler OutlineDrawing gibi "SECTION &lt;kesim&gt;" katmanında, grup başına bir etiket, başlık
    /// 400 boyunda altta. Hepsi OTO_AILE ile işaretlenir; sonraki çalıştırmada silinip yeniden çizilir.</summary>
    internal static class UzatmaCizimleri
    {
        internal const string Uygulama = "OTO_AILE";
        /// <summary>Son çizilen uzatma kesitlerinin adları (K2, L3 ...): tabloda olmamaları beklenir.</summary>
        internal static readonly HashSet<string> SonAdlar = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        internal static string Ciz(Form f, bool stil)
        {
            Document doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return null;
            Database db = doc.Database;
            string yol = FormBaglayici.YolAl(f);
            if (yol.Length == 0 || !File.Exists(yol)) return null;
            List<AileParcasi> parcalar = AileParcalari.Oku(TowYama.SatirlarHam(yol));
            if (parcalar.Count == 0)
            {
                // aile dosyası değil: önceki bir aile dosyasından kalan uzatma çizimleri silinir
                using (doc.LockDocument())
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    var ms = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);
                    foreach (ObjectId id in ms)
                    {
                        var e = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (e != null && Bizim(e)) { try { e.UpgradeOpen(); e.Erase(); } catch { } }
                    }
                    tr.Commit();
                }
                SonAdlar.Clear();
                return null;
            }
            List<UzatmaCizimi> cizimler = UzatmaPlani.Hesapla(parcalar);
            SonAdlar.Clear();
            foreach (UzatmaCizimi c in cizimler) if (c.Kesit) SonAdlar.Add(c.Baslik.Substring("SECTION ".Length));
            KuleVerisi kule = KuleOkuyucu.Oku(f.GetType().Assembly, yol);
            var redundant = kule != null ? kule.RedundantGruplar : new HashSet<string>();

            int adet = 0;
            using (doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var ms = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);
                var rat = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
                if (!rat.Has(Uygulama))
                {
                    rat.UpgradeOpen();
                    var r = new RegAppTableRecord { Name = Uygulama };
                    rat.Add(r);
                    tr.AddNewlyCreatedDBObject(r, true);
                }
                KesitDizici.UygulamaKaydi(db, tr);
                string grupAdi = null;   // çizilen nesneler bu adla da işaretlenir (ölçü / dizme kendi çizimine bağlar)
                // önceki çizimler silinir; mevcut çizimin sınırı ve OutlineDrawing etiket boyu bulunur
                double xMax = double.MinValue, zMin = double.MaxValue, w = 0;
                var boylar = new List<double>();
                bool xz = true;
                foreach (ObjectId id in ms)
                {
                    var e = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (e == null) continue;
                    if (Bizim(e)) { try { e.UpgradeOpen(); e.Erase(); } catch { } continue; }
                    var t = e as DBText;
                    if (t != null)
                    {
                        if (t.TextString != null && t.TextString.Trim().StartsWith("SECTION ", StringComparison.OrdinalIgnoreCase) && t.Height >= 250)
                        { xz = Math.Abs(t.Normal.Y) > 0.9; w = xz ? t.Position.Y : t.Position.Z; }
                        else if (t.Height > 20 && t.Height < 350) boylar.Add(t.Height);
                    }
                    try
                    {
                        Extents3d x = e.GeometricExtents;
                        xMax = Math.Max(xMax, x.MaxPoint.X);
                        zMin = Math.Min(zMin, xz ? x.MinPoint.Z : x.MinPoint.Y);
                    }
                    catch { }
                }
                if (xMax == double.MinValue) { xMax = 0; zMin = 0; }
                boylar.Sort();
                double h = boylar.Count > 0 ? boylar[boylar.Count / 2] : 150;
                Func<double, double, Point3d> P = (u, v) => xz ? new Point3d(u, w, v) : new Point3d(u, v, w);
                Matrix3d duzlem = xz ? Matrix3d.Displacement(new Vector3d(0, w, 0)) * Matrix3d.Rotation(Math.PI / 2, Vector3d.XAxis, Point3d.Origin)
                                     : Matrix3d.Displacement(new Vector3d(0, 0, w));
                ObjectId kesik = stil ? KesitStili.KesikCizgiTipi(db, tr) : ObjectId.Null;
                ObjectId duz = SymbolUtilityServices.GetLinetypeContinuousId(db);
                double olcek = stil ? KesitStili.DesenOlcegi(db) : 1;
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                Func<string, ObjectId> katman = ad =>
                {
                    if (!lt.Has(ad))
                    {
                        if (!lt.IsWriteEnabled) lt.UpgradeOpen();
                        var ltr = new LayerTableRecord { Name = ad };
                        lt.Add(ltr);
                        tr.AddNewlyCreatedDBObject(ltr, true);
                    }
                    return lt[ad];
                };
                Action<Entity> ekle = e =>
                {
                    ms.AppendEntity(e); tr.AddNewlyCreatedDBObject(e, true);
                    e.XData = new ResultBuffer(new TypedValue((int)DxfCode.ExtendedDataRegAppName, Uygulama),
                                               new TypedValue((int)DxfCode.ExtendedDataAsciiString, "1"));
                    if (grupAdi != null)
                        e.XData = new ResultBuffer(new TypedValue((int)DxfCode.ExtendedDataRegAppName, KesitDizici.Uygulama),
                                                   new TypedValue((int)DxfCode.ExtendedDataAsciiString, grupAdi));
                };
                Action<string, double, double, double, double> yazi = (metin, u, v, aci, boy) =>
                {
                    var t = new DBText();
                    t.SetDatabaseDefaults(db);
                    t.Layer = "0";
                    t.TextString = metin; t.Height = boy; t.Rotation = aci;
                    t.HorizontalMode = TextHorizontalMode.TextCenter; t.VerticalMode = TextVerticalMode.TextBase;
                    t.AlignmentPoint = new Point3d(u, v, 0);
                    t.TransformBy(duzlem);
                    ekle(t);
                    try { t.AdjustAlignment(db); } catch { }
                };

                double ox = xMax + 10000, aralik = 3000;
                foreach (UzatmaCizimi c in cizimler)
                {
                    double u0 = c.Cizgiler.Min(q => Math.Min(q[0], q[2])), u1 = c.Cizgiler.Max(q => Math.Max(q[0], q[2]));
                    double v0 = c.Cizgiler.Min(q => Math.Min(q[1], q[3]));
                    double du = ox - u0, dv = zMin - v0;
                    grupAdi = c.Kesit ? c.Baslik.Substring("SECTION ".Length) : c.Baslik;
                    var enUzun = new Dictionary<string, int>();
                    for (int i = 0; i < c.Cizgiler.Count; i++)
                    {
                        double[] q = c.Cizgiler[i];
                        var l = new Line(P(q[0] + du, q[1] + dv), P(q[2] + du, q[3] + dv));
                        l.SetDatabaseDefaults(db);
                        l.LayerId = katman("SECTION " + (string.IsNullOrEmpty(c.Kesimler[i]) ? "0" : c.Kesimler[i]));
                        if (stil) { try { KesitStili.Boya(l, redundant.Contains((c.Gruplar[i] ?? "").Trim()), kesik, duz, olcek); } catch { } }
                        ekle(l);
                        adet++;
                        string g = c.Gruplar[i] ?? "";
                        int j;
                        if (g.Length > 0 && (!enUzun.TryGetValue(g, out j) || Boy(c.Cizgiler[j]) < Boy(q))) enUzun[g] = i;
                    }
                    foreach (var kv in enUzun)
                    {
                        double[] q = c.Cizgiler[kv.Value];
                        double a = Math.Atan2(q[3] - q[1], q[2] - q[0]);
                        if (a > Math.PI / 2 + 1e-6) a -= Math.PI; else if (a <= -Math.PI / 2 + 1e-6) a += Math.PI;
                        double mu = (q[0] + q[2]) / 2 + du - Math.Sin(a) * 0.3 * h, mv = (q[1] + q[3]) / 2 + dv + Math.Cos(a) * 0.3 * h;
                        yazi(kv.Key, mu, mv, a, h);
                    }
                    yazi(c.Baslik, (u0 + u1) / 2 + du, zMin - 1200, 0, 400);
                    ox += (u1 - u0) + aralik;
                }
                tr.Commit();
            }
            return cizimler.Count == 0 ? null : string.Format("Aile dosyası: {0} uzatma çizimi ({1})", cizimler.Count,
                string.Join(", ", cizimler.Select(c => c.Baslik.Replace("SECTION ", ""))));
        }

        private static double Boy(double[] q) { return Math.Sqrt((q[2] - q[0]) * (q[2] - q[0]) + (q[3] - q[1]) * (q[3] - q[1])); }

        private static bool Bizim(Entity e)
        {
            try { using (ResultBuffer rb = e.GetXDataForApplication(Uygulama)) return rb != null; }
            catch { return false; }
        }
    }

    /// <summary>
    /// Eğik düzlem kesitlerini çizimci gibi döndürür. OutlineDrawing düzlem kesitini, düzlemin yatay doğrultusu sağa ve
    /// eğim (yukarı) doğrultusu yukarı gelecek şekilde çizer; travers üst düzlemi gibi eğimi kulenin X ekseni boyunca
    /// olan düzlemlerde travers dikey çıkar (ucu aşağıda). Burada düzlemdeki X'e en yakın doğrultu yatay ve +X sağda
    /// olacak şekilde kesit 90° döndürülür (görünüşteki ve TAO paftasındaki gibi travers yatay). Yazılar okunur kalır,
    /// başlık yine altta ortalanır; döndürülen kesitin nesneleri kesit adıyla işaretlenir (ölçü / dizme onları tanır).
    /// </summary>
    internal static class KesitYonu
    {
        internal const string Uygulama = "OTO_KESIT_YON";

        /// <summary>Derece, + saat yönü tersi; 0: döndürme gerekmez (yatay kesit, 45°'den dik düzlem, köşegen bacak düzlemi).</summary>
        internal static int Aci(KesitTanimi t)
        {
            double[] u, v;
            if (t == null || t.Yukseklik.HasValue || !KesitSecici.Eksenler(t, out u, out v)) return 0;
            // dik düzlem (45°'den dik: yüz, çerçeve) görünüş gibi dik kalır; plana yakın düzlem plan gibi okunur
            if (v[2] > Math.Sqrt(0.5)) return 0;
            if (Math.Abs(v[0]) <= Math.Abs(u[0]) + 0.05) return 0;
            return v[0] < 0 ? 90 : -90;   // yukarısı -X'e bakıyorsa saat yönü tersi: yukarı sola, +X sağa gelir
        }

        internal static string Dondur(Form f)
        {
            Document doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return null;
            Database db = doc.Database;
            var acilar = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in KesitStili.GridTanimlari(f)) { int a = Aci(kv.Value); if (a != 0) acilar[kv.Key] = a; }
            if (acilar.Count == 0) return null;

            var donen = new List<string>();
            using (doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var ms = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);
                KesitDizici.Gruplar g = KesitDizici.Grupla(tr, ms);
                if (g.Kesitler.Count == 0) { tr.Commit(); return null; }
                KesitDizici.UygulamaKaydi(db, tr);
                var rat = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
                if (!rat.Has(Uygulama))
                {
                    rat.UpgradeOpen();
                    var r = new RegAppTableRecord { Name = Uygulama };
                    rat.Add(r);
                    tr.AddNewlyCreatedDBObject(r, true);
                }
                Vector3d eksen = g.XZ ? -Vector3d.YAxis : Vector3d.ZAxis;   // ekrana dik, bakana doğru: + açı saat yönü tersi
                Func<double, double, double, Point3d> P = (u, v, w) => g.XZ ? new Point3d(u, w, v) : new Point3d(u, v, w);
                Func<double, double, Vector3d> V = (du, dv) => g.XZ ? new Vector3d(du, 0, dv) : new Vector3d(du, dv, 0);

                foreach (KesitDizici.Baslik b in g.Kesitler)
                {
                    int aci;
                    if (!acilar.TryGetValue(b.Ad, out aci)) continue;
                    var baslik = tr.GetObject(b.Id, OpenMode.ForRead) as Entity;
                    if (baslik == null || Donduruldu(baslik)) continue;
                    List<ObjectId> ids = g.Nesneler[b.Ad].Where(id => id != b.Id).ToList();
                    double[] k = null;
                    foreach (ObjectId id in ids) k = KesitDizici.Birlestir(k, KesitDizici.Kutu((Entity)tr.GetObject(id, OpenMode.ForRead), g.XZ));
                    double[] bk = KesitDizici.Kutu(baslik, g.XZ);
                    if (k == null || bk == null) continue;

                    Matrix3d M = Matrix3d.Rotation(aci * Math.PI / 180.0, eksen, P((k[0] + k[2]) / 2, (k[1] + k[3]) / 2, b.W));
                    foreach (ObjectId id in ids)
                    {
                        var e = (Entity)tr.GetObject(id, OpenMode.ForWrite);
                        e.TransformBy(M);
                        var t = e as DBText;
                        if (t != null) Okunur(t, g.XZ, eksen);
                    }
                    // başlık: yeni kutunun altında, eski boşlukla, ortalı
                    double[] k2 = null;
                    foreach (ObjectId id in ids) k2 = KesitDizici.Birlestir(k2, KesitDizici.Kutu((Entity)tr.GetObject(id, OpenMode.ForRead), g.XZ));
                    if (k2 == null) k2 = k;
                    double bosluk = Math.Max(0, k[1] - bk[3]);
                    baslik.UpgradeOpen();
                    baslik.TransformBy(Matrix3d.Displacement(V((k2[0] + k2[2]) / 2 - (bk[0] + bk[2]) / 2, k2[1] - bosluk - bk[3])));
                    ids.Add(b.Id);
                    KesitDizici.Tasi(tr, ids, b.Ad, new Vector3d(0, 0, 0));
                    baslik.XData = new ResultBuffer(new TypedValue((int)DxfCode.ExtendedDataRegAppName, Uygulama),
                                                    new TypedValue((int)DxfCode.ExtendedDataInteger16, (short)aci));
                    donen.Add(b.Ad + (aci > 0 ? " ↺" : " ↻"));
                }
                tr.Commit();
            }
            return donen.Count == 0 ? null : "Kesit yönü: " + string.Join(", ", donen) + " 90° döndürüldü (travers yatay)";
        }

        private static bool Donduruldu(Entity e)
        {
            try { using (ResultBuffer rb = e.GetXDataForApplication(Uygulama)) return rb != null; }
            catch { return false; }
        }

        /// <summary>Döndürmeden sonra baş aşağı / sağdan sola okunan yazıyı kendi merkezinde 180° çevirir.</summary>
        private static void Okunur(DBText t, bool xz, Vector3d eksen)
        {
            try
            {
                Vector3d d = Vector3d.XAxis.RotateBy(t.Rotation, Vector3d.ZAxis).TransformBy(Matrix3d.PlaneToWorld(t.Normal));
                double du = d.X, dv = xz ? d.Z : d.Y;
                if (du > 1e-6 || (Math.Abs(du) <= 1e-6 && dv > 0)) return;
                Extents3d x = t.GeometricExtents;
                Point3d m = new Point3d((x.MinPoint.X + x.MaxPoint.X) / 2, (x.MinPoint.Y + x.MaxPoint.Y) / 2, (x.MinPoint.Z + x.MaxPoint.Z) / 2);
                t.TransformBy(Matrix3d.Rotation(Math.PI, eksen, m));
            }
            catch { }
        }
    }

    /// <summary>Ön görünüşteki OutlineDrawing kesit işaretlerini (SECTION_ARROW_ blokları) KesitIsaretPlani'na göre yenileriyle
    /// değiştirir. İşaretler OutlineDrawing'in katmanında (SECTION_MARK): kesit gruplama / pafta onları görünüşe sayar.</summary>
    internal static class KesitIsaretleri
    {
        internal const string Katman = "SECTION_MARK";

        internal static string Ciz(Form f)
        {
            Document doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return null;
            Database db = doc.Database;
            string yol = FormBaglayici.YolAl(f);
            if (yol.Length == 0 || !File.Exists(yol)) return null;
            KuleVerisi kule = KuleOkuyucu.Oku(f.GetType().Assembly, yol);
            Dictionary<string, KesitTanimi> tanimlar = KesitStili.GridTanimlari(f);
            if (kule == null || tanimlar.Count == 0) return null;

            int adet = 0, silinen = 0;
            using (doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var ms = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);
                GorunusKonumu on = GorunusYeri.Bul(tr, ms, kule).FirstOrDefault(k => !k.Yan);
                if (on == null) { tr.Commit(); return null; }

                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (!lt.Has(Katman))
                {
                    lt.UpgradeOpen();
                    var ltr = new LayerTableRecord { Name = Katman };
                    lt.Add(ltr);
                    tr.AddNewlyCreatedDBObject(ltr, true);
                }
                ObjectId katman = lt[Katman];

                // OutlineDrawing'in okları ve önceki çalıştırmanın işaretleri
                foreach (ObjectId id in ms)
                {
                    var e = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (e == null) continue;
                    bool sil = false;
                    var br = e as BlockReference;
                    if (br != null)
                    {
                        var btr = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                        sil = btr != null && btr.Name.StartsWith("SECTION_ARROW_", StringComparison.OrdinalIgnoreCase);
                    }
                    else if ((e is Line || e is DBText) && string.Equals(e.Layer, Katman, StringComparison.OrdinalIgnoreCase)) sil = true;
                    if (!sil) continue;
                    try { e.UpgradeOpen(); e.Erase(); silinen++; } catch { }
                }

                Func<double, double, Point3d> N = (u, v) => on.IkiD ? new Point3d(u + on.Ofs, v, 0) : new Point3d(u + on.Ofs, 0, v);
                Matrix3d duzlem = on.IkiD ? Matrix3d.Identity : Matrix3d.Rotation(Math.PI / 2, Vector3d.XAxis, Point3d.Origin);
                foreach (KesitIsareti s in KesitIsaretPlani.Hesapla(kule.OnYuz, tanimlar))
                {
                    foreach (double[] c in s.Cizgiler)
                    {
                        var l = new Line(N(c[0], c[1]), N(c[2], c[3]));
                        l.SetDatabaseDefaults(db); l.LayerId = katman;
                        ms.AppendEntity(l); tr.AddNewlyCreatedDBObject(l, true);
                    }
                    var t = new DBText();
                    t.SetDatabaseDefaults(db); t.LayerId = katman;
                    t.TextString = s.Ad; t.Height = s.YaziBoyu;
                    t.HorizontalMode = TextHorizontalMode.TextCenter; t.VerticalMode = TextVerticalMode.TextBase;
                    t.AlignmentPoint = new Point3d(s.YaziU + on.Ofs, s.YaziV, 0);
                    t.TransformBy(duzlem);
                    ms.AppendEntity(t); tr.AddNewlyCreatedDBObject(t, true);
                    try { t.AdjustAlignment(db); } catch { }
                    adet++;
                }
                tr.Commit();
            }
            return string.Format("Kesit işaretleri: {0} ({1} OutlineDrawing oku kaldırıldı)", adet, silinen);
        }
    }

    /// <summary>
    /// Görünüşte (ön: bakış +Y, yan: bakış -X yönü, OutlineDrawing'in sideFace dönüşümü (x, y) -> (y, -x)) yüz düzleminde
    /// olmayan elemanların izdüşümü. PLS-TOWER görünüşü kulenin tamamını izdüşürür; OutlineDrawing yalnız yüz düzlemini çizer.
    /// İzdüşümü bir noktaya düşen (bakış yönündeki), yüz çizgilerinin veya daha önce eklenen bir izdüşümün üstüne binen
    /// (aynı doğru üzerinde, onlarca kapsanan) elemanlar eklenmez; arka yüz, yan yüzler ve yatay düzlemler böylece elenir,
    /// kalça çaprazları ve iç diyaframlar kalır. Ayna kopyalardan yalnız biri (aynı izdüşüm) eklenir.
    /// </summary>
    public static class IcElemanPlani
    {
        public static double EnKisa = 20;        // bundan kısa izdüşüm (bakış yönündeki eleman) eklenmez (mm)
        public static double MesafeTol = 5;      // aynı doğru üzerinde sayılma (mm)
        public static double AciTol = 0.005;     // ~0.3°
        /// <summary>İzdüşüm boyu / gerçek boy bundan küçükse (eleman bakış yönüne yakın; ör. yan görünüşte
        /// uçtan görülen travers elemanları) eklenmez; yoksa kol uçtan bakışta okunmaz bir yığın olur.</summary>
        public static double EnAzOran = 0.3;
        public static double DerinlikTol = 50;   // gövde içi sayılma payı (mm)

        /// <returns>görünüş koordinatlarında (ofset hariç) çizgiler: P = { u, derinlik, z, u, derinlik, z }</returns>
        public static List<YuzCizgisi> Hesapla(IDictionary<string, FinalMember> uyeler, List<YuzCizgisi> yuz, bool yan)
        {
            var sonuc = new List<YuzCizgisi>();
            var yuzdeki = new HashSet<string>(yuz.Where(c => c.Anahtar != null).Select(c => c.Anahtar));
            var mevcut = yuz.Select(c => new[] { c.P[0], c.P[2], c.P[3], c.P[5] }).ToList();
            var adaylar = new List<YuzCizgisi>();
            foreach (var kv in uyeler)
            {
                if (yuzdeki.Contains(kv.Key)) continue;
                FinalMember m = kv.Value;
                double[] p = yan
                    ? new[] { m.start_y, -m.start_x, m.start_z, m.end_y, -m.end_x, m.end_z }
                    : new[] { m.start_x, m.start_y, m.start_z, m.end_x, m.end_y, m.end_z };
                double izd = Boy(p[0], p[2], p[3], p[5]);
                double gercek = Math.Sqrt(izd * izd + (p[4] - p[1]) * (p[4] - p[1]));
                if (izd < EnKisa) continue;   // bakış yönünde: izdüşümü nokta
                // Bakış yönüne yakın veya gövdenin dışında (bakış yönünde yüz düzleminden öteye taşan) elemanlar:
                // uçtan görülen travers / tepe. Eksik çizgi olmasın diye bunlar da çizilir, ama ayrı katmanda (Uc).
                bool uc = izd < EnAzOran * gercek ||
                          Math.Abs(p[1]) > Derinlik(yuz, p[2]) + DerinlikTol || Math.Abs(p[4]) > Derinlik(yuz, p[5]) + DerinlikTol;
                adaylar.Add(new YuzCizgisi { P = p, Grup = m.group_label, Kesim = m.section_label, Anahtar = kv.Key, Uc = uc });
            }
            // Önce gövde içi, sonra uçtan görülenler; her grupta uzun izdüşümler önce: kısa olanlar onların üstüne
            // biniyorsa elenir. Eşit boyda önde (bakana yakın) olan.
            foreach (var c in adaylar.OrderBy(c => c.Uc).ThenByDescending(c => Math.Round(Boy(c.P[0], c.P[2], c.P[3], c.P[5])))
                                     .ThenBy(c => Math.Min(c.P[1], c.P[4])).ThenBy(c => c.Anahtar, StringComparer.Ordinal))
            {
                double[] s = { c.P[0], c.P[2], c.P[3], c.P[5] };
                if (Kapsaniyor(s, mevcut)) continue;
                mevcut.Add(s);
                sonuc.Add(c);
            }
            return sonuc;
        }

        public static double BoslukPayi = 3000;  // yüzün o kotta çizgisi yoksa en yakın yüz kotu bu kadar uzakta olabilir (mm)

        /// <summary>z kotunda görünüş yüzünün bakış yönündeki en büyük |derinliği| (gövde yarı genişliği).
        /// Yüzün o kotta çizgisi yoksa (ör. OutlineDrawing'in yan yüzü, önünde uçtan görülen travers olan en üst paneli
        /// atlıyor) en yakın yüz kotundaki değer; o da BoslukPayi'ndan uzaksa 0.</summary>
        private static double Derinlik(List<YuzCizgisi> yuz, double z)
        {
            double d = 0, enYakin = double.MaxValue, yakinD = 0;
            bool var = false;
            foreach (var c in yuz)
            {
                double z1 = c.P[2], z2 = c.P[5];
                if (z < Math.Min(z1, z2) - 1 || z > Math.Max(z1, z2) + 1)
                {
                    for (int i = 0; i < 2; i++)
                    {
                        double dz = Math.Abs(c.P[2 + 3 * i] - z), dd = Math.Abs(c.P[1 + 3 * i]);
                        if (dz < enYakin - 1) { enYakin = dz; yakinD = dd; }
                        else if (Math.Abs(dz - enYakin) <= 1) yakinD = Math.Max(yakinD, dd);
                    }
                    continue;
                }
                var = true;
                double v = Math.Abs(z2 - z1) < 1 ? Math.Max(Math.Abs(c.P[1]), Math.Abs(c.P[4]))
                                                 : Math.Abs(c.P[1] + (c.P[4] - c.P[1]) * (z - z1) / (z2 - z1));
                d = Math.Max(d, v);
            }
            if (var) return d;
            return enYakin <= BoslukPayi ? yakinD : 0;
        }

        private static double Boy(double u1, double v1, double u2, double v2)
        {
            return Math.Sqrt((u2 - u1) * (u2 - u1) + (v2 - v1) * (v2 - v1));
        }

        /// <summary>s, kendisiyle aynı doğru üzerindeki çizgilerin birleşimi tarafından tamamen örtülüyor mu?</summary>
        public static bool Kapsaniyor(double[] s, List<double[]> cizgiler)
        {
            double du = s[2] - s[0], dv = s[3] - s[1], L = Math.Sqrt(du * du + dv * dv);
            if (L < 1e-9) return true;
            double eu = du / L, ev = dv / L;
            var araliklar = new List<double[]>();
            foreach (var c in cizgiler)
            {
                double cu = c[2] - c[0], cv = c[3] - c[1], cl = Math.Sqrt(cu * cu + cv * cv);
                if (cl < 1e-9 || Math.Abs(eu * cv - ev * cu) / cl > AciTol) continue;
                // c'nin uçlarının s doğrusuna uzaklığı
                if (Math.Abs((c[0] - s[0]) * ev - (c[1] - s[1]) * eu) > MesafeTol) continue;
                if (Math.Abs((c[2] - s[0]) * ev - (c[3] - s[1]) * eu) > MesafeTol) continue;
                double t1 = (c[0] - s[0]) * eu + (c[1] - s[1]) * ev, t2 = (c[2] - s[0]) * eu + (c[3] - s[1]) * ev;
                double a = Math.Min(t1, t2), b = Math.Max(t1, t2);
                if (b < -MesafeTol || a > L + MesafeTol) continue;
                araliklar.Add(new[] { a, b });
            }
            double kapali = 0;
            foreach (var r in araliklar.OrderBy(r => r[0]))
            {
                if (r[0] > kapali + MesafeTol) return false;
                kapali = Math.Max(kapali, r[1]);
                if (kapali >= L - MesafeTol) return true;
            }
            return kapali >= L - MesafeTol;
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
            var zincir = new List<OlcuTanimi>();
            for (int i = 0; i + 1 < kotlar.Count; i++)
                zincir.Add(t(new OlcuTanimi
                {
                    U1 = Sag(segs, kotlar[i]), V1 = kotlar[i], U2 = Sag(segs, kotlar[i + 1]), V2 = kotlar[i + 1],
                    UD = uz, VD = (kotlar[i] + kotlar[i + 1]) / 2, Rotation = Math.PI / 2
                }));
            DuseyZincirYerlestir(zincir, uz);
            sonuc.AddRange(zincir);
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
                if (solUc) sonuc.Add(TekYerlestir(t(new OlcuTanimi { U1 = ul, V1 = v, U2 = eksen, V2 = v, UD = (ul + eksen) / 2, VD = vd, Rotation = 0 }), false));
                if (sagUc) sonuc.Add(TekYerlestir(t(new OlcuTanimi { U1 = eksen, V1 = v, U2 = ur, V2 = v, UD = (ur + eksen) / 2, VD = vd, Rotation = 0 }), true));
            }
            return sonuc;
        }

        /// <summary>Yatay tek ölçü (uç): yazı sığmıyorsa eksenden uzak tarafa, uzatma çizgisinin dışına.</summary>
        private static OlcuTanimi TekYerlestir(OlcuTanimi o, bool saga)
        {
            if (OlcuPlani.YaziSigar(o)) return o;
            double w = OlcuPlani.YaziGenisligi(o);
            o.YaziU = saga ? Math.Max(o.U1, o.U2) + 0.75 * o.Yazi + w / 2 : Math.Min(o.U1, o.U2) - 0.75 * o.Yazi - w / 2;
            o.YaziV = o.VD + 0.5 * o.Yazi;
            return o;
        }

        /// <summary>Düşey zincirde yazısı sığmayan kısa parçalar: en alttaki aşağıya, en üstteki yukarıya, aradakiler ölçü
        /// çizgisinin dış (sağ) yanına.</summary>
        private static void DuseyZincirYerlestir(List<OlcuTanimi> zincir, double uz)
        {
            for (int i = 0; i < zincir.Count; i++)
            {
                OlcuTanimi o = zincir[i];
                if (OlcuPlani.YaziSigar(o)) continue;
                double w = OlcuPlani.YaziGenisligi(o), alt = Math.Min(o.V1, o.V2), ust = Math.Max(o.V1, o.V2);
                if (i == 0 && zincir.Count > 1) { o.YaziU = uz - 0.5 * o.Yazi; o.YaziV = alt - 0.75 * o.Yazi - w / 2; }
                else if (i == zincir.Count - 1 && zincir.Count > 1) { o.YaziU = uz - 0.5 * o.Yazi; o.YaziV = ust + 0.75 * o.Yazi + w / 2; }
                else if (zincir.Count > 1) { o.YaziU = uz + 0.75 * o.Yazi; o.YaziV = (alt + ust) / 2; }
            }
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
                || IcElemanlar.BizimKatman(katman)
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
    /// <summary>
    /// Tek DLL yaması: OutlineDrawing.dataprocessing.IsPoseLine yerine. Düğüm / eleman tablosunda sonraki kaydın ad
    /// satırı mı? Asıl kod adın 15 karakterden kısa ve boşluksuz olmasını istiyordu; uzun bir ad tablonun geri kalanını
    /// sessizce düşürüyordu. Burada: boş değil, ';' yok (tablo sonundaki PLS başlık satırlarının hepsinde ';' var) ve
    /// tek ad (tırnaklı ad boşluk içerebilir).
    /// </summary>
    public static class TowYama
    {
        /// <summary>Tek DLL yaması: runtowerfile'da görünüşler arası ofset (asıl kodda sabit 50000).</summary>
        public static int GorunusAraligi(object dataprocessing) { return Aralik(Genislik(dataprocessing), 50000, 10000); }

        /// <summary>Tek DLL yaması: runtowerfile'da kesitler arası ofset (asıl kodda sabit 30000).</summary>
        public static int KesitAraligi(object dataprocessing) { return Aralik(Genislik(dataprocessing), 30000, 6000); }

        /// <summary>En az asgari; kule yatay genişliği + pay'dan küçükse 1000'e yuvarlanmış genişlik + pay.</summary>
        public static int Aralik(double genislik, int asgari, int pay)
        {
            if (double.IsNaN(genislik) || genislik <= 0) return asgari;
            return Math.Max(asgari, (int)(Math.Ceiling((genislik + pay) / 1000.0) * 1000));
        }

        /// <summary>finalMember'ların x ve y aralığından büyüğü (mm); okunamazsa 0.</summary>
        public static double Genislik(object dataprocessing)
        {
            try
            {
                const BindingFlags H = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
                var fm = dataprocessing.GetType().GetProperty("finalMember", H).GetValue(dataprocessing, null) as IDictionary;
                if (fm == null || fm.Count == 0) return 0;
                PropertyInfo[] k = null;
                double x0 = double.MaxValue, x1 = double.MinValue, y0 = double.MaxValue, y1 = double.MinValue;
                foreach (DictionaryEntry de in fm)
                {
                    if (de.Value == null) continue;
                    if (k == null) k = new[] { "start_x", "start_y", "end_x", "end_y" }.Select(a => de.Value.GetType().GetProperty(a, H)).ToArray();
                    double sx = Convert.ToDouble(k[0].GetValue(de.Value, null)), sy = Convert.ToDouble(k[1].GetValue(de.Value, null));
                    double ex = Convert.ToDouble(k[2].GetValue(de.Value, null)), ey = Convert.ToDouble(k[3].GetValue(de.Value, null));
                    x0 = Math.Min(x0, Math.Min(sx, ex)); x1 = Math.Max(x1, Math.Max(sx, ex));
                    y0 = Math.Min(y0, Math.Min(sy, ey)); y1 = Math.Max(y1, Math.Max(sy, ey));
                }
                return x0 > x1 ? 0 : Math.Max(x1 - x0, y1 - y0);
            }
            catch { return 0; }
        }

        /// <summary>Eklenti tarafı: kule verisinden aynı hesap.</summary>
        public static double Genislik(IDictionary<string, FinalMember> uyeler)
        {
            if (uyeler == null || uyeler.Count == 0) return 0;
            double x0 = uyeler.Values.Min(m => Math.Min(m.start_x, m.end_x)), x1 = uyeler.Values.Max(m => Math.Max(m.start_x, m.end_x));
            double y0 = uyeler.Values.Min(m => Math.Min(m.start_y, m.end_y)), y1 = uyeler.Values.Max(m => Math.Max(m.start_y, m.end_y));
            return Math.Max(x1 - x0, y1 - y0);
        }

        /// <summary>Tek DLL yaması: datapreprocessing'deki File.ReadAllLines(yol) yerine. Asıl kod UTF-8 çözüyordu; PLS
        /// dosyayı Windows-1254 yazar ve Ç, Ş, Ğ, İ, Ö, Ü, ı hepsi aynı U+FFFD'ye iniyordu: yalnız bu harflerde farklı
        /// iki etiket tek anahtara düşüp biri sessizce kayboluyordu. Geçerli UTF-8 ise (ASCII dahil) UTF-8, değilse
        /// 1254 (yoksa bayt = karakter, 28591); satırlara bölme File.ReadAllLines ile aynı.</summary>
        public static string[] Satirlar(string yol)
        {
            return TowAilesi.TekKonfig(SatirlarHam(yol), TowAilesi.Secili);
        }

        /// <summary>Satirlar'ın aile süzmesi yapılmamış hali (bütün konfigürasyonlar).</summary>
        public static string[] SatirlarHam(string yol)
        {
            byte[] b = File.ReadAllBytes(yol);
            string metin;
            int bom = b.Length >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF ? 3 : 0;
            try { metin = new System.Text.UTF8Encoding(false, true).GetString(b, bom, b.Length - bom); }
            catch (ArgumentException)
            {
                System.Text.Encoding k;
                try { k = System.Text.Encoding.GetEncoding(1254); } catch { k = System.Text.Encoding.GetEncoding(28591); }
                metin = k.GetString(b);
            }
            var sonuc = new List<string>();
            using (var r = new StringReader(metin))
            {
                string satir;
                while ((satir = r.ReadLine()) != null) sonuc.Add(satir);
            }
            return sonuc.ToArray();
        }

        /// <summary>Tek DLL yaması: SectionDetection sonundaki dump(towYolu) yerine. .tow klasörüne tow_section_dump.txt
        /// yazılamazsa (salt okunur ağ klasörü, Excel'de açık dosya) istisna bütün kesitlerin çizimini iptal ediyordu.</summary>
        public static void DumpGuvenli(object dataprocessing, string towYolu)
        {
            try
            {
                MethodInfo d = dataprocessing.GetType().GetMethod("dump", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                    null, new[] { typeof(string) }, null);
                if (d != null) d.Invoke(dataprocessing, new object[] { towYolu });
            }
            catch { /* döküm yalnız bilgi amaçlı */ }
        }

        /// <summary>Tek DLL yaması: Parse_SectionLabel'de "tam 15 alan" yerine. Kayıt sayısı başlıktan ("9 ; section
        /// label ..."); alan sayısı en az 15 ve renk alanı tamsayı (daha yeni / eski PLS sürümünde alan sayısı
        /// değişince bütün kesit katmanları oluşmuyor, ön / yan görünüş hiç çizilmiyordu).</summary>
        public static bool KesitEtiketSatiri(string[] satirlar, int i, int baslik, string[] alanlar)
        {
            int n, renk;
            if (BaslikSayisi(satirlar, baslik, out n) && i - baslik - 1 >= n) return false;
            return alanlar != null && alanlar.Length >= 15 && alanlar[0] != null && alanlar[2] != null
                && int.TryParse(alanlar[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out renk);
        }

        /// <summary>Tek DLL yaması: düğüm / eleman tablosu okuma döngüsünde i'den başlayan kayıt var mı. Tablo sonu
        /// başlıktaki kayıt sayısından ("128 ; Joints Geometry ...") anlaşılır; sayı okunamazsa PozSatiri.</summary>
        public static bool KayitVar(string[] satirlar, int i, int baslik, int adim)
        {
            if (satirlar == null || i < 0 || i >= satirlar.Length) return false;
            int n;
            if (BaslikSayisi(satirlar, baslik, out n))
            {
                if ((i - baslik - 1) / Math.Max(1, adim) >= n || i + adim - 1 >= satirlar.Length) return false;
                // sayı yanlışsa sonraki tablonun başlığına taşma
                return !string.IsNullOrWhiteSpace(satirlar[i]) && satirlar[i].IndexOf(';') < 0;
            }
            return PozSatiri(satirlar[i]);
        }

        public static bool BaslikSayisi(string[] satirlar, int baslik, out int n)
        {
            n = 0;
            if (satirlar == null || baslik < 0 || baslik >= satirlar.Length || satirlar[baslik] == null) return false;
            string b = satirlar[baslik];
            int nk = b.IndexOf(';');
            if (nk <= 0) return false;
            return int.TryParse(b.Substring(0, nk).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out n) && n >= 0;
        }

        public static bool PozSatiri(string satir)
        {
            if (string.IsNullOrWhiteSpace(satir) || satir.IndexOf(';') >= 0) return false;
            string s = satir.Trim();
            if (s[0] == '\'' || s[0] == '"')
                return s.Length >= 2 && s.IndexOf(s[0], 1) == s.Length - 1 && s.Substring(1, s.Length - 2).Trim().Length > 0;
            return s.IndexOfAny(new[] { ' ', '\t' }) < 0 && !YolGibi(s);
        }

        /// <summary>Tablodan sonra gelen bolt kütüphanesi (.blt) vb. dosya yolu satırı.</summary>
        public static bool YolGibi(string s)
        {
            return s.IndexOf('\\') >= 0 || s.IndexOf('/') >= 0 || s.IndexOf(':') >= 0
                || s.EndsWith(".blt", StringComparison.OrdinalIgnoreCase);
        }
    }

    public static class KesitAnahtari
    {
        public const char Ayrac = '\u00A7';   // §

        public static string Olustur(string uyeAnahtari, string kesitAdi)
        {
            return (kesitAdi ?? "") + Ayrac + uyeAnahtari;
        }
    }

    /// <summary>
    /// PLS-TOWER "family" dosyası: gövdenin ardından her gövde uzatması × bacak uzatması birleşimi ayrı bir konfigürasyon
    /// olarak gelir ('TAO_12BE_4.0L,...' adı, destekler, gövde uzatması bloğu, dört bacak bloğu). Hepsi aynı koordinat
    /// sistemindedir; OutlineDrawing hepsini birden okuyunca uzatmalar üst üste biner, aynı etiketli elemanlar birbirinin
    /// üstüne yazılır ve görünüşün alt kısmı karmakarışık olur. Burada OutlineDrawing'e yalnız bir konfigürasyon verilir
    /// (varsayılan ilki: dosyadaki en uzun, çizimcinin ana görünüşü).
    /// </summary>
    public static class TowAilesi
    {
        /// <summary>OutlineDrawing'e verilecek konfigürasyonun sırası (0: dosyadaki ilk).</summary>
        public static int Secili;

        /// <summary>Konfigürasyon adı satırlarının indeksleri: tırnakla başlayan ve ardından "; support label" satırı gelen satır.</summary>
        public static List<int> Baslar(string[] s)
        {
            var l = new List<int>();
            for (int i = 0; i + 1 < s.Length; i++)
                if (s[i].StartsWith("'", StringComparison.Ordinal) && s[i + 1].IndexOf("; support label", StringComparison.Ordinal) >= 0)
                    l.Add(i);
            return l;
        }

        public static string Ad(string satir)
        {
            int j = satir.IndexOf('\'', 1);
            return j > 1 ? satir.Substring(1, j - 1) : satir;
        }

        public static List<string> Adlar(string[] s)
        {
            return Baslar(s).Select(i => Ad(s[i])).ToList();
        }

        /// <summary>Aile dosyası değilse (en fazla bir konfigürasyon) satırlar olduğu gibi döner. Aileyse yalnız baş kısım:
        /// PLS onu varsayılan tam model olarak yazar (gövde + bir gövde uzatması + dört bacak, bacak elemanları LX_/LY_
        /// önekleriyle açılmış). Konfigürasyon bloklarındaki düğümler ekli adlarla ('100P', '115X') yazıldığı ve
        /// OutlineDrawing her düğüm adına "P" eklediği için bloklar OutlineDrawing'e verilmez; diğer uzatmalar
        /// AileParcalari ile eklenti tarafında okunur.</summary>
        public static string[] TekKonfig(string[] s, int k)
        {
            List<int> b = Baslar(s);
            if (b.Count < 2) return s;
            var l = new List<string>(b[0]);
            for (int i = 0; i < b[0]; i++) l.Add(s[i]);
            if (l.Count > 0)
            {
                // "25 0 22.259 ..." (konfigürasyon sayısı) -> "0 ..."
                string say = l[l.Count - 1].TrimStart();
                int bosluk = say.IndexOf(' ');
                int n;
                if (bosluk > 0 && int.TryParse(say.Substring(0, bosluk), out n) && n == b.Count) l[l.Count - 1] = "0" + say.Substring(bosluk);
            }
            return l.ToArray();
        }
    }

    /// <summary>Aile dosyasındaki bir parçanın (gövde ya da bacak uzatması) elemanı: uçlar mm, grup ve kesim etiketi.</summary>
    public class ParcaElemani
    {
        public double[] A, B;
        public string Grup, Kesim;
    }

    /// <summary>Aile dosyasında bir gövde ya da bacak uzatması bloğu.</summary>
    public class AileParcasi
    {
        public string Ad;            // "+9.0 Bodt Ext." (dosyadaki gibi)
        public bool Bacak;
        public double Uzatma;        // +9, +3 ...
        public string Konfig;        // ilk geçtiği konfigürasyon
        public List<ParcaElemani> Elemanlar = new List<ParcaElemani>();
    }

    /// <summary>
    /// Aile dosyasındaki uzatma bloklarını OutlineDrawing'e gerek kalmadan okur. Bloklarda düğümler açılmış adlarla
    /// ('100P', '115S', '115X', '100XY') ve koordinatlarıyla, elemanlar da açılmış olarak (g52P, g52X ...) yazılır: simetri
    /// uygulanmaz. Bacak bloklarında PLS düğümleri yerel tanımlar (en uzun gövde uzatmasına göre); başlıktaki kaydırma
    /// (dz, dx, dy) eklenir. Her uzatma adı için ilk geçtiği blok alınır; bacak bloklarından yalnız ilki (dördü aynı,
    /// +X+Y köşesinde tanımlı).
    /// </summary>
    public static class AileParcalari
    {
        private static readonly Regex Baslik = new Regex(@"^'([^']*(Bod[a-z]*|Leg) Ext\.?)'\s+(.*)$", RegexOptions.IgnoreCase);
        private static readonly Regex Dugum = new Regex(@"^\s*(\d+)\s*;\s*(Joints Geometry|Secondary Joints)");
        private static readonly Regex Eleman = new Regex(@"^\s*(\d+)\s*;\s*Angle Member Connectivity");
        private static readonly Regex GrupKesim = new Regex(@"^'([^']*)'\s+'([^']*)'\s*;\s*group, section");
        private static readonly Regex Sayi = new Regex(@"[-+]?\d+(\.\d+)?");

        /// <summary>Aile dosyası değilse boş liste. Sıra: dosyadaki ilk geçiş sırası.</summary>
        public static List<AileParcasi> Oku(string[] s)
        {
            var sonuc = new List<AileParcasi>();
            List<int> b = TowAilesi.Baslar(s);
            if (b.Count < 2) return sonuc;
            var ci = CultureInfo.InvariantCulture;
            var gorulen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int k = 0; k < b.Count; k++)
            {
                int son = k + 1 < b.Count ? b[k + 1] : s.Length;
                string konfig = TowAilesi.Ad(s[b[k]]);
                // bloklar: başlık satırından bir sonraki başlığa kadar
                var basliklar = new List<int>();
                for (int i = b[k]; i < son; i++) if (Baslik.IsMatch(s[i])) basliklar.Add(i);
                for (int q = 0; q < basliklar.Count; q++)
                {
                    Match mb = Baslik.Match(s[basliklar[q]]);
                    string ad = mb.Groups[1].Value.Trim();
                    if (!gorulen.Add(ad)) continue;
                    int bs = basliklar[q], bson = q + 1 < basliklar.Count ? basliklar[q + 1] : son;
                    bool bacak = mb.Groups[2].Value.Equals("Leg", StringComparison.OrdinalIgnoreCase);
                    string[] t = mb.Groups[3].Value.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                    double dx = 0, dy = 0, dz = 0;
                    if (bacak && t.Length >= 8)
                    {
                        double.TryParse(t[3], NumberStyles.Float, ci, out dz);
                        double.TryParse(t[6], NumberStyles.Float, ci, out dx);
                        double.TryParse(t[7], NumberStyles.Float, ci, out dy);
                    }
                    Match mu = Sayi.Match(ad);
                    double uz = 0;
                    if (mu.Success) double.TryParse(mu.Value, NumberStyles.Float, ci, out uz);
                    var p = new AileParcasi { Ad = ad, Bacak = bacak, Uzatma = uz, Konfig = konfig };
                    BlokOku(s, bs, bson, dx, dy, dz, p.Elemanlar);
                    if (p.Elemanlar.Count > 0) sonuc.Add(p);
                }
            }
            return sonuc;
        }

        private static void BlokOku(string[] s, int bas, int son, double dx, double dy, double dz, List<ParcaElemani> l)
        {
            var ci = CultureInfo.InvariantCulture;
            var dugum = new Dictionary<string, double[]>();
            for (int i = bas; i < son; i++)
            {
                Match md = Dugum.Match(s[i]);
                if (md.Success)
                {
                    int n = int.Parse(md.Groups[1].Value, ci);
                    for (int r = 0; r < n && i + 1 + 5 * r + 2 < son; r++)
                    {
                        int j = i + 1 + 5 * r;
                        string[] t = s[j + 2].Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                        double x, y, z;
                        if (t.Length < 3 || !double.TryParse(t[0], NumberStyles.Float, ci, out x) || !double.TryParse(t[1], NumberStyles.Float, ci, out y) ||
                            !double.TryParse(t[2], NumberStyles.Float, ci, out z)) continue;
                        dugum[s[j].Trim().Trim('\'')] = new[] { (x + dx) * 1000, (y + dy) * 1000, (z + dz) * 1000 };
                    }
                    i += 5 * n;
                    continue;
                }
                Match me = Eleman.Match(s[i]);
                if (!me.Success) continue;
                int m = int.Parse(me.Groups[1].Value, ci);
                for (int r = 0; r < m && i + 1 + 8 * r + 4 < son; r++)
                {
                    int j = i + 1 + 8 * r;
                    double[] a, b;
                    if (!dugum.TryGetValue(s[j + 1].Trim().Trim('\''), out a) || !dugum.TryGetValue(s[j + 2].Trim().Trim('\''), out b)) continue;
                    Match mg = GrupKesim.Match(s[j + 4]);
                    l.Add(new ParcaElemani { A = a, B = b, Grup = mg.Success ? mg.Groups[1].Value : "", Kesim = mg.Success ? mg.Groups[2].Value : "" });
                }
                i += 8 * m;
            }
        }
    }

    /// <summary>Uzatma çizimi: görünüş ("+9 BODY EXTENSION") ya da kesit ("SECTION K2"); düzlem koordinatlarında.</summary>
    public class UzatmaCizimi
    {
        public string Baslik;
        public bool Kesit;
        public List<double[]> Cizgiler = new List<double[]>();   // { u1, v1, u2, v2 }
        public List<string> Gruplar = new List<string>();         // çizgiyle aynı sırada
        public List<string> Kesimler = new List<string>();
    }

    /// <summary>
    /// Çizimcinin aile paftasındaki uzatma çizimleri (220kV TAO örneği): ana görünüşte olmayan her gövde uzatması için
    /// ön yüz görünüşü ve alt planı (K kesitleri), her bacak uzatması için ön yüz görünüşü ve iç düzlemi (L kesitleri:
    /// ayak ile iki kasık düğümünden geçen düzlem, ters üçgen). Ana görünüşteki uzatmaların (baş kısımdaki varsayılan
    /// model) kesitleri otomatik kesitlerden gelir; K / L numaraları onlardan sonra başlar.
    /// </summary>
    public static class UzatmaPlani
    {
        public static List<UzatmaCizimi> Hesapla(List<AileParcasi> parcalar)
        {
            var sonuc = new List<UzatmaCizimi>();
            if (parcalar == null || parcalar.Count == 0) return sonuc;
            // varsayılan (ilk konfigürasyonun) parçaları ana görünüşte
            string ilk = parcalar[0].Konfig;
            var govdeler = parcalar.Where(p => !p.Bacak && p.Konfig != ilk).ToList();
            var bacaklar = parcalar.Where(p => p.Bacak && p.Konfig != ilk).ToList();
            var ci = CultureInfo.InvariantCulture;
            int k = 2;
            foreach (var p in govdeler)
            {
                string u = p.Uzatma.ToString("0.#", ci);
                sonuc.Add(Yuz(p, "+" + u + " BODY EXTENSION"));
                UzatmaCizimi plan = AltPlan(p, "SECTION K" + k++);
                if (plan != null) sonuc.Add(plan);
            }
            int l = 2;
            foreach (var p in bacaklar)
            {
                string u = p.Uzatma.ToString("0.#", ci);
                sonuc.Add(Yuz(p, "+" + u + " LEG EXTENSION"));
                UzatmaCizimi ic = IcDuzlem(p, "SECTION L" + l++);
                if (ic != null) sonuc.Add(ic);
            }
            return sonuc.Where(c => c.Cizgiler.Count > 0).ToList();
        }

        private static void Ekle(UzatmaCizimi c, ParcaElemani e, double u1, double v1, double u2, double v2)
        {
            if (Math.Abs(u1 - u2) < 1 && Math.Abs(v1 - v2) < 1) return;
            c.Cizgiler.Add(new[] { u1, v1, u2, v2 });
            c.Gruplar.Add(e.Grup);
            c.Kesimler.Add(e.Kesim);
        }

        /// <summary>Ön yüz (+Y yüzü): köşe hattına göre y(z) doğrusu; iki ucu da bu düzlemde olan elemanlar (x, z).</summary>
        private static UzatmaCizimi Yuz(AileParcasi p, string baslik)
        {
            var c = new UzatmaCizimi { Baslik = baslik };
            var uclar = p.Elemanlar.SelectMany(e => new[] { e.A, e.B }).ToList();
            double z0 = uclar.Min(q => q[2]), z1 = uclar.Max(q => q[2]);
            double yAlt = uclar.Where(q => q[2] < z0 + 1).Max(q => q[1]), yUst = uclar.Where(q => q[2] > z1 - 1).Max(q => q[1]);
            Func<double[], bool> yuzde = q =>
            {
                double yc = z1 - z0 < 1 ? yUst : yAlt + (yUst - yAlt) * (q[2] - z0) / (z1 - z0);
                return Math.Abs(q[1] - yc) < 30;
            };
            foreach (var e in p.Elemanlar)
                if (yuzde(e.A) && yuzde(e.B)) Ekle(c, e, e.A[0], e.A[2], e.B[0], e.B[2]);
            return c;
        }

        /// <summary>Gövde uzatmasının alt kotundaki yatay elemanlar, plan (x, y).</summary>
        private static UzatmaCizimi AltPlan(AileParcasi p, string baslik)
        {
            double z0 = p.Elemanlar.Min(e => Math.Min(e.A[2], e.B[2]));
            var c = new UzatmaCizimi { Baslik = baslik, Kesit = true };
            foreach (var e in p.Elemanlar)
                if (Math.Abs(e.A[2] - z0) < 15 && Math.Abs(e.B[2] - z0) < 15) Ekle(c, e, e.A[0], e.A[1], e.B[0], e.B[1]);
            return c.Cizgiler.Count >= 3 ? c : null;
        }

        /// <summary>Bacağın iç düzlemi: ayak (en düşük, x+y en büyük) ile üst kottaki iki kasık düğümünden (biri en küçük x,
        /// biri en küçük y) geçen düzlem; iki ucu düzleme 50 mm'den yakın elemanlar. Eksenler OutlineDrawing gibi: u
        /// düzlemin yatay doğrultusu (+X'e), v yukarı.</summary>
        private static UzatmaCizimi IcDuzlem(AileParcasi p, string baslik)
        {
            var uclar = p.Elemanlar.SelectMany(e => new[] { e.A, e.B }).ToList();
            double z0 = uclar.Min(q => q[2]), z1 = uclar.Max(q => q[2]);
            double[] ayak = uclar.Where(q => q[2] < z0 + 1).OrderByDescending(q => q[0] + q[1]).First();
            var ust = uclar.Where(q => q[2] > z1 - 1).ToList();
            double[] k1 = ust.OrderBy(q => q[0]).First(), k2 = ust.OrderBy(q => q[1]).First();
            double[] a = Fark(k1, ayak), b = Fark(k2, ayak);
            double[] n = Birim(Capraz(a, b));
            if (n == null) return null;
            double d = Nokta(n, ayak);
            double[] u = Birim(Fark(k2, k1));
            if (u == null) return null;
            if (u[0] < 0) u = new[] { -u[0], -u[1], -u[2] };
            double[] v = Birim(Capraz(n, u));
            if (v[2] < 0) v = new[] { -v[0], -v[1], -v[2] };
            var c = new UzatmaCizimi { Baslik = baslik, Kesit = true };
            foreach (var e in p.Elemanlar)
            {
                if (Math.Abs(Nokta(n, e.A) - d) > 50 || Math.Abs(Nokta(n, e.B) - d) > 50) continue;
                Ekle(c, e, Nokta(u, e.A), Nokta(v, e.A), Nokta(u, e.B), Nokta(v, e.B));
            }
            return c.Cizgiler.Count >= 3 ? c : null;
        }

        private static double[] Fark(double[] a, double[] b) { return new[] { a[0] - b[0], a[1] - b[1], a[2] - b[2] }; }
        private static double Nokta(double[] a, double[] b) { return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]; }
        private static double[] Capraz(double[] a, double[] b) { return new[] { a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2], a[0] * b[1] - a[1] * b[0] }; }
        private static double[] Birim(double[] a) { double l = Math.Sqrt(Nokta(a, a)); return l < 1e-9 ? null : new[] { a[0] / l, a[1] / l, a[2] / l }; }
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
        /// <param name="okunan">OutlineDrawing'in okuduğu eleman anahtarları (finalMember); verilirse .tow'daki elemanlarla
        /// karşılaştırılır ve okunmayan eleman varsa uyarılır.</param>
        private static string _uyariAnahtar, _uyariSonuc;

        /// <summary>Aynı .tow (yol + değişme zamanı + boyut) ve aynı okunan eleman sayısı için son sonuç döner.</summary>
        public static string Uyari(string towPath, ICollection<string> okunan = null)
        {
            string anahtar = null;
            try
            {
                var fi = new FileInfo(towPath);
                if (fi.Exists) anahtar = fi.FullName + "|" + fi.LastWriteTimeUtc.Ticks + "|" + fi.Length + "|" + (okunan == null ? -1 : okunan.Count);
            }
            catch { }
            if (anahtar != null && anahtar == _uyariAnahtar) return _uyariSonuc;
            string sonuc = UyariHesapla(towPath, okunan);
            if (anahtar != null) { _uyariAnahtar = anahtar; _uyariSonuc = sonuc; }
            return sonuc;
        }

        private static string UyariHesapla(string towPath, ICollection<string> okunan)
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
                    Komutlar.BirKezYaz(string.Format("\nBilgi: simetri kodu {0} olan {1} elemanın aynaları eklendi (OutlineDrawing düzeltmesi).\n", kod, l.Count));
                    d.Remove(kod);
                }
#endif
                var mesajlar = new List<string>();
                if (d.Count > 0)
                {
                    int toplam = d.Values.Sum(l => l.Count);
                    string kodlar = string.Join(", ", d.Keys.OrderBy(x => x));
                    string ornek = string.Join(", ", d.Values.SelectMany(l => l).Take(4));
                    Komutlar.BirKezYaz(string.Format("\nUYARI: .tow'da simetri kodu {0} olan {1} eleman var; OutlineDrawing bunların aynalarını çizmiyor " +
                        "(çizimde eksik eleman olur). Elemanlar: {2}\n", kodlar, toplam,
                        string.Join(", ", d.Values.SelectMany(l => l))));
                    mesajlar.Add(string.Format("DİKKAT: {0} elemanın simetri kodu {1}; aynaları çizilmiyor ({2}...)", toplam, kodlar, ornek));
                }
                string[] satirlar = TowYama.Satirlar(towPath);   // OutlineDrawing ile aynı kod çözme (tek DLL yaması)
                string m;
                if (okunan != null && (m = OkumaKontrol(satirlar, okunan)) != null) mesajlar.Add(m);
                if ((m = Birim(satirlar)) != null) mesajlar.Add(m);
                if ((m = DigerElemanlar(satirlar)) != null) mesajlar.Add(m);
                if ((m = DugumuOlmayanlar(satirlar)) != null) mesajlar.Add(m);
                if ((m = GrupTablosu(satirlar)) != null) mesajlar.Add(m);
                return mesajlar.Count == 0 ? null : string.Join(" | ", mesajlar);
            }
            catch { return null; }
        }

        /// <summary>.tow'daki açı elemanlarını (aynalarıyla) kendisi sayar, OutlineDrawing'in okuduklarıyla karşılaştırır.
        /// Eksik varsa (okuyucunun bilinmeyen bir sınırı, bozuk kayıt ...) adlarıyla uyarır: çizimden detay düşmesin.</summary>
        public static string OkumaKontrol(string[] satirlar, ICollection<string> okunan)
        {
            List<string> beklenen = BeklenenUyeler(satirlar);
            if (beklenen.Count == 0) return null;
            var okunanSet = new HashSet<string>(okunan, StringComparer.Ordinal);
            var eksik = beklenen.Where(k => !okunanSet.Contains(k)).ToList();
            if (eksik.Count == 0) return null;
            Komutlar.BirKezYaz(string.Format("\nUYARI: .tow'da {0} açı elemanı (aynalarıyla) var, OutlineDrawing {1} tanesini okudu. ÇİZİLMEYECEK {2} eleman: {3}\n",
                beklenen.Count, beklenen.Count - eksik.Count, eksik.Count, string.Join(", ", eksik.Take(60)) + (eksik.Count > 60 ? " ..." : "")));
            return string.Format("DİKKAT: {0} eleman OutlineDrawing tarafından okunmadı, çizilmeyecek ({1}...)", eksik.Count, string.Join(", ", eksik.Take(3)));
        }

        /// <summary>Düğümleri tanımlı açı elemanlarının OutlineDrawing anahtarları: ad + "" / X / Y / XY (simetri koduna göre).</summary>
        public static List<string> BeklenenUyeler(string[] satirlar)
        {
            var dugum = new HashSet<string>(StringComparer.Ordinal);
            var sonuc = new List<string>();
            var bol = new[] { ' ', '\t' };
            for (int i = 0; i < satirlar.Length; i++)
            {
                string s = satirlar[i];
                int yorum = s.IndexOf(';'), n;
                if (yorum < 0 || !int.TryParse(s.Substring(0, yorum).Trim(), out n)) continue;
                bool ikincil = s.IndexOf("Secondary Joints", yorum, StringComparison.Ordinal) >= 0;
                if (ikincil || s.IndexOf("Joints Geometry", yorum, StringComparison.Ordinal) >= 0)
                {
                    for (int j = 0, k = i + 1; j < n && k + 3 < satirlar.Length; j++, k += 5)
                    {
                        string ad = satirlar[k].Trim().Trim('\'', '"').Trim();
                        string[] t = satirlar[k + 3].Split(bol, StringSplitOptions.RemoveEmptyEntries);
                        int kod = t.Length > 1 && int.TryParse(t[1], out kod) ? kod : 0;
                        dugum.Add(ad + (ikincil ? "S" : "P"));
                        if (kod == 1 || kod == 3) dugum.Add(ad + "X");
                        if (kod == 2 || kod == 3) dugum.Add(ad + "Y");
                        if (kod == 3) dugum.Add(ad + "XY");
                    }
                }
                else if (s.IndexOf("Angle Member Connectivity", yorum, StringComparison.Ordinal) >= 0)
                {
                    for (int j = 0, k = i + 1; j < n && k + 5 < satirlar.Length; j++, k += 8)
                    {
                        string ad = satirlar[k].Trim();
                        if (!dugum.Contains(satirlar[k + 1].Trim()) || !dugum.Contains(satirlar[k + 2].Trim())) continue;   // DugumuOlmayanlar bildirir
                        string[] t = satirlar[k + 5].Split(bol, StringSplitOptions.RemoveEmptyEntries);
                        int kod = t.Length > 0 && int.TryParse(t[0], out kod) ? kod : 0;
                        sonuc.Add(ad);
#if TEK_DLL
                        if (kod == 1 || kod == 3 || kod == 12) sonuc.Add(ad + "X");   // 12: SimetriDuzeltici
#else
                        if (kod == 1 || kod == 3) sonuc.Add(ad + "X");
#endif
                        if (kod == 2 || kod == 3) sonuc.Add(ad + "Y");
                        if (kod == 3) sonuc.Add(ad + "XY");
                    }
                }
            }
            return sonuc;
        }

        private static readonly Regex BirimDeseni = new Regex(@"UNITS\s*=\s*'([^']*)'", RegexOptions.IgnoreCase);

        /// <summary>PLS-TOWER .tow'u iç birimle (UNITS='INTERNAL', metre) yazar; OutlineDrawing koordinatları metre kabul edip
        /// 1000 ile çarpar. Başka bir birim yazıyorsa çizim yanlış ölçekte çıkar.</summary>
        public static string Birim(string[] satirlar)
        {
            if (satirlar.Length == 0) return null;
            Match mt = BirimDeseni.Match(satirlar[0]);
            if (!mt.Success) return null;
            string b = mt.Groups[1].Value.Trim();
            if (b.Equals("INTERNAL", StringComparison.OrdinalIgnoreCase) || b.Equals("SI", StringComparison.OrdinalIgnoreCase) ||
                b.Equals("METRIC", StringComparison.OrdinalIgnoreCase)) return null;
            Komutlar.BirKezYaz("\nUYARI: .tow birimi UNITS='" + b + "'. OutlineDrawing koordinatları metre kabul ediyor; çizim yanlış ölçekte olabilir.\n");
            return "DİKKAT: .tow birimi '" + b + "' (metre bekleniyor)";
        }

        private static readonly Regex DigerTur = new Regex(@"^\s*(\d+)\s*;\s*((?:Structure Cable|Guy|Brace|Davit Arm|Tubular Davit Arm|X-Arm|Tubular X-Arm) Connectivity)",
            RegexOptions.IgnoreCase);

        /// <summary>OutlineDrawing yalnız açı elemanlarını (Angle Member) okur; kablo, gergi, davit, X-arm gibi elemanlar çizilmez.</summary>
        public static string DigerElemanlar(string[] satirlar)
        {
            var l = new List<string>();
            int toplam = 0;
            foreach (string s in satirlar)
            {
                Match mt = DigerTur.Match(s);
                int n;
                if (!mt.Success || !int.TryParse(mt.Groups[1].Value, out n) || n <= 0) continue;
                l.Add(n + " " + mt.Groups[2].Value);
                toplam += n;
            }
            if (toplam == 0) return null;
            Komutlar.BirKezYaz("\nUYARI: .tow'da açı elemanı olmayan elemanlar var (" + string.Join(", ", l) +
                "). OutlineDrawing yalnız açı elemanlarını okur; bunlar görünüşlerde ve kesitlerde OLMAYACAK.\n");
            return "DİKKAT: " + toplam + " eleman (kablo/davit/X-arm...) açı elemanı değil, çizilmez";
        }

        /// <summary>Düğümü tanımlı olmayan açı elemanları (OutlineDrawing bunları sessizce atlar).</summary>
        public static string DugumuOlmayanlar(string[] satirlar)
        {
            var dugum = new HashSet<string>(StringComparer.Ordinal);
            var eksik = new List<string>();
            var bol = new[] { ' ', '\t' };
            for (int i = 0; i < satirlar.Length; i++)
            {
                string s = satirlar[i];
                int yorum = s.IndexOf(';'), n;
                if (yorum < 0 || !int.TryParse(s.Substring(0, yorum).Trim(), out n)) continue;
                bool ikincil = s.IndexOf("Secondary Joints", yorum, StringComparison.Ordinal) >= 0;
                if (ikincil || s.IndexOf("Joints Geometry", yorum, StringComparison.Ordinal) >= 0)
                {
                    for (int j = 0, k = i + 1; j < n && k + 3 < satirlar.Length; j++, k += 5)
                    {
                        string ad = satirlar[k].Trim().Trim('\'', '"').Trim();
                        string[] t = satirlar[k + 3].Split(bol, StringSplitOptions.RemoveEmptyEntries);
                        int kod = t.Length > 1 && int.TryParse(t[1], out kod) ? kod : 0;
                        dugum.Add(ad + (ikincil ? "S" : "P"));
                        if (kod == 1 || kod == 3) dugum.Add(ad + "X");
                        if (kod == 2 || kod == 3) dugum.Add(ad + "Y");
                        if (kod == 3) dugum.Add(ad + "XY");
                    }
                }
                else if (s.IndexOf("Angle Member Connectivity", yorum, StringComparison.Ordinal) >= 0)
                {
                    for (int j = 0, k = i + 1; j < n && k + 2 < satirlar.Length; j++, k += 8)
                    {
                        string a = satirlar[k + 1].Trim(), b = satirlar[k + 2].Trim();
                        if (!dugum.Contains(a) || !dugum.Contains(b)) eksik.Add(satirlar[k].Trim());
                    }
                }
            }
            if (eksik.Count == 0) return null;
            Komutlar.BirKezYaz("\nUYARI: düğümü tanımlı olmayan " + eksik.Count + " açı elemanı var; OutlineDrawing bunları çizmez: " + string.Join(", ", eksik) + "\n");
            return "DİKKAT: " + eksik.Count + " elemanın düğümü yok, çizilmez (" + string.Join(", ", eksik.Take(3)) + "...)";
        }

        /// <summary>Grup tablosunda 9'dan az alanlı satır: OutlineDrawing grup okumayı orada bırakır; sonraki grupların
        /// profil / malzeme bilgisi etiketlerde çıkmaz. (Fazla alan tek DLL'de sorun değil.)</summary>
        public static string GrupTablosu(string[] satirlar)
        {
            for (int i = 0; i < satirlar.Length; i++)
            {
                string s = satirlar[i];
                int yorum = s.IndexOf(';'), n;
                if (yorum < 0 || s.IndexOf("group label", yorum, StringComparison.OrdinalIgnoreCase) < 0 ||
                    !int.TryParse(s.Substring(0, yorum).Trim(), out n)) continue;
                for (int k = i + 1; k <= i + n && k < satirlar.Length; k++)
                {
                    string govde = satirlar[k];
                    int y = govde.IndexOf(';');
                    if (y >= 0) govde = govde.Substring(0, y);
                    int alan = Regex.Matches(govde, "'[^']*'|\\S+").Count;
#if TEK_DLL
                    bool sorun = alan < 9;
#else
                    bool sorun = alan != 9;
#endif
                    if (!sorun) continue;
                    int kalan = i + n - k + 1;
                    Komutlar.BirKezYaz(string.Format("\nUYARI: grup tablosunun {0}. satırında {1} alan var (OutlineDrawing 9 bekliyor): '{2}'. " +
                        "OutlineDrawing grup okumayı burada bırakır; bu ve sonraki {3} grubun profil / malzeme bilgisi etiketlerde çıkmaz.\n",
                        k - i, alan, satirlar[k].Trim(), kalan));
                    return "DİKKAT: grup tablosu " + (k - i) + ". satırdan sonra okunamıyor (" + kalan + " grup etiketsiz)";
                }
                return null;
            }
            return null;
        }
    }

    /// <summary>OutlineDrawing'in .tow okumasından gelen elemanlar ve ön/yan görünüşte görünen eleman anahtarları.</summary>
    public class KuleVerisi
    {
        /// <summary>Kule yatay genişliği (mm).</summary>
        public double Genislik { get { return TowYama.Genislik(Uyeler); } }
        /// <summary>Kule yüksekliği (mm).</summary>
        public double Yukseklik
        {
            get
            {
                if (Uyeler == null || Uyeler.Count == 0) return 0;
                return Uyeler.Values.Max(m => Math.Max(m.start_z, m.end_z)) - Uyeler.Values.Min(m => Math.Min(m.start_z, m.end_z));
            }
        }
        /// <summary>OutlineDrawing'in görünüşler arası ofseti (tek DLL'de kule genişliğine göre, ayrı eklentide 50000).</summary>
        public int GorunusAraligi
        {
#if TEK_DLL
            get { return TowYama.Aralik(Genislik, 50000, 10000); }
#else
            get { return 50000; }
#endif
        }
        /// <summary>OutlineDrawing'in kesitler arası ofseti (tek DLL'de kule genişliğine göre, ayrı eklentide 30000).</summary>
        public int KesitAraligi
        {
#if TEK_DLL
            get { return TowYama.Aralik(Genislik, 30000, 6000); }
#else
            get { return 30000; }
#endif
        }
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
        public string Anahtar;  // finalMember anahtarı (g338, g338X ...)
        public bool Uc;         // iç eleman izdüşümü: bakış yönüne yakın / gövde dışında (uçtan görülen travers)
    }

    /// <summary>Ön/yan görünüş çizgilerini koordinatından tanır. OutlineDrawing görünüşü
    /// Point3d(x + offset, y, z) ile çizer (ön: offset 0, yan: ön de çizildiyse 50000);
    /// "2D Aktar" sonrası hali: y = 0'a basılıp X ekseni etrafında -90° döndürülmüştür: (x, z, 0).
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
                    if (ikiBoyutlu) Ekle(IkiD(ax, c.P[1], c.P[2], bx, c.P[4], c.P[5]), c.Grup);   // 2D Aktar sonrası
                }
        }

        /// <summary>OutlineDrawing'in 2D Aktar'ı (main.Converto2D): önce FlattenY0 (y = 0), sonra X ekseni etrafında -90°
        /// döndürme -> (x, y, z) çizgisi (x, z, 0) olur. İki ucu z = 0'da olan çizgi (IsEntityOnXY) dönüştürülmez.</summary>
        public static double[] IkiD(double ax, double ay, double az, double bx, double by, double bz)
        {
            if (Math.Abs(az) <= 1e-6 && Math.Abs(bz) <= 1e-6) return new[] { ax, ay, 0.0, bx, by, 0.0 };
            return new[] { ax, az, 0.0, bx, bz, 0.0 };
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

        /// <summary>Yalnız OutlineDrawing'in Run'ı null döndüğünde true (asıl okuyucu çöktü); hata = sebebi.</summary>
        public static bool Okunamiyor(Assembly outlineDrawing, string towPath, out string hata)
        {
            hata = null;
            if (Oku(outlineDrawing, towPath) != null) return false;
            hata = "bilinmeyen okuma hatası";
            try
            {
                Type dp = outlineDrawing.GetType("OutlineDrawing.dataprocessing", true);
                MethodInfo on = dp.GetMethod("datapreprocessing", Hepsi, null, new[] { typeof(string) }, null);
                if (on != null) on.Invoke(Activator.CreateInstance(dp, true), new object[] { towPath });
            }
            catch (System.Exception ex)
            {
                System.Exception ic = ex;
                while (ic is TargetInvocationException && ic.InnerException != null) ic = ic.InnerException;
                string yer = ic.StackTrace != null ? ic.StackTrace.Split('\n').Select(x => x.Trim())
                    .FirstOrDefault(x => x.Contains("Parse_") || x.Contains("dataprocessing")) : null;
                hata = ic.GetType().Name + ": " + ic.Message + (yer != null ? " (" + yer + ")" : "");
            }
            return true;
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
                    Kesim = ks != null ? ks.GetValue(v, null) as string : null,
                    Anahtar = Convert.ToString(de.Key, CultureInfo.InvariantCulture)
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
            // OutlineDrawing ile aynı kod çözme (TowYama.Satirlar; tek DLL'de datapreprocessing de bunu kullanır):
            // Türkçe karakterli grup adları iki tarafta da aynı biçimde çözülsün ki anahtarlar eşleşsin.
            string[] satirlar = TowYama.Satirlar(towPath);
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
