// Uçtan uca detay testi (AutoCAD'siz): "hiçbir eleman atlanmıyor mu?"
//  1. .tow bağımsız okunur (bu dosyadaki okuyucu) ve OutlineDrawing'in okuduğu elemanlarla geometri olarak
//     karşılaştırılır: OutlineDrawing'in düşürdüğü (eksik düğüm, desteklenmeyen simetri) eleman var mı?
//     Açı elemanı dışındaki eleman türleri (kablo, davit, X-arm, ...) sayılır: OutlineDrawing bunları çizmez.
//  2. Eklentinin otomatik kesitleri bulunur, OutlineDrawing'in GERÇEK SectionDetection'ı çalıştırılır.
//  3. Her eleman ailesi (birbirinin X/Y aynası olan elemanlar) için: ön görünüşte, yan görünüşte veya bir kesitte
//     kendi çizgisiyle (ve OutlineDrawing etiketiyle) çiziliyor mu; değilse iç eleman izdüşümü olarak mı görünüyor;
//     hiç mi görünmüyor?
//   bash test/detay_testi.sh eklenti/bin/tek/OutlineDrawing.dll kule.tow [ayrinti]
using System; using System.Collections; using System.Collections.Generic; using System.Globalization;
using System.IO; using System.Linq; using System.Reflection; using System.Text.RegularExpressions; using OtomatikKesit;

public static class DetayTesti
{
    const BindingFlags H = BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;

    // ---------------- bağımsız .tow okuyucu ----------------
    class Uye { public string Ad, Grup; public double[] P; }

    static List<Uye> TowOku(string yol, out Dictionary<string, int> digerTurler, out int eksikDugum)
    {
        string[] L = File.ReadAllText(yol, System.Text.Encoding.GetEncoding(28591)).Split('\n').Select(s => s.TrimEnd('\r')).ToArray();
        var dugum = new Dictionary<string, double[]>();
        var uyeler = new List<Uye>();
        digerTurler = new Dictionary<string, int>();
        eksikDugum = 0;
        var ham = new List<Tuple<string, string, string, string, int>>();
        for (int i = 0; i < L.Length; i++)
        {
            string s = L[i];
            int yorum = s.IndexOf(';');
            if (yorum < 0) continue;
            int n;
            if (!int.TryParse(s.Substring(0, yorum).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out n)) continue;
            string baslik = s.Substring(yorum + 1);
            if (baslik.Contains("Joints Geometry") || baslik.Contains("Secondary Joints"))
            {
                string kendi = baslik.Contains("Secondary") ? "S" : "P";
                int k = i + 1;
                for (int j = 0; j < n; j++, k += 5)
                {
                    string ad = L[k].Trim().Trim('\'');
                    double[] xyz = L[k + 2].Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries).Take(3)
                                    .Select(v => double.Parse(v, CultureInfo.InvariantCulture)).ToArray();
                    int sym = int.Parse(L[k + 3].Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries)[1], CultureInfo.InvariantCulture);
                    dugum[ad + kendi] = xyz;
                    // Ayna düğümler: kod 1 -> X (y -> -y), 2 -> Y (x -> -x), 3 -> X, Y, XY
                    if (sym == 1 || sym == 3) dugum[ad + "X"] = new[] { xyz[0], -xyz[1], xyz[2] };
                    if (sym == 2 || sym == 3) dugum[ad + "Y"] = new[] { -xyz[0], xyz[1], xyz[2] };
                    if (sym == 3) dugum[ad + "XY"] = new[] { -xyz[0], -xyz[1], xyz[2] };
                }
                i = k - 1;
            }
            else if (baslik.Contains("Angle Member Connectivity"))
            {
                int k = i + 1;
                for (int j = 0; j < n; j++, k += 8)
                {
                    var g = Regex.Matches(L[k + 4], "'([^']*)'").Cast<Match>().Select(m => m.Groups[1].Value.Trim()).ToList();
                    int sym = int.Parse(L[k + 5].Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries)[0], CultureInfo.InvariantCulture);
                    ham.Add(Tuple.Create(L[k].Trim(), L[k + 1].Trim(), L[k + 2].Trim(), g.Count > 0 ? g[0] : "", sym));
                }
                i = k - 1;
            }
            else if (n > 0 && Regex.IsMatch(baslik, @"(Cable|Guy|Brace|Davit|X-Arm|Equipment|CAN) Connectivity"))
                digerTurler[baslik.Trim().Split(':')[0]] = n;
        }
        foreach (var h in ham)
        {
            double[] a, b;
            if (!dugum.TryGetValue(h.Item2, out a) || !dugum.TryGetValue(h.Item3, out b)) { eksikDugum++; continue; }
            var v = new List<Tuple<string, double, double>> { Tuple.Create("", 1.0, 1.0) };
            int sym = h.Item5;
            if (sym == 1 || sym == 3 || sym == 12) v.Add(Tuple.Create("X", 1.0, -1.0));   // 12: tek DLL düzeltmesi, yalnız X
            if (sym == 2 || sym == 3) v.Add(Tuple.Create("Y", -1.0, 1.0));
            if (sym == 3) v.Add(Tuple.Create("XY", -1.0, -1.0));
            foreach (var t in v)
                uyeler.Add(new Uye { Ad = h.Item1 + t.Item1, Grup = h.Item4,
                    P = new[] { a[0] * t.Item2, a[1] * t.Item3, a[2], b[0] * t.Item2, b[1] * t.Item3, b[2] } });
        }
        // OutlineDrawing ile aynı: m -> mm, (x, y) -> (y, -x), en düşük z < 0 ise 0'a kaydır
        double zmin = uyeler.Count == 0 ? 0 : uyeler.Min(u => Math.Min(u.P[2], u.P[5]));
        if (zmin >= -1e-9) zmin = 0;
        foreach (var u in uyeler)
            u.P = new[] { u.P[1] * 1000, -u.P[0] * 1000, (u.P[2] - zmin) * 1000, u.P[4] * 1000, -u.P[3] * 1000, (u.P[5] - zmin) * 1000 };
        return uyeler;
    }

    static string Anahtar(double[] p)
    {
        Func<double, string> R = v => (Math.Round(v) + 0.0).ToString(CultureInfo.InvariantCulture);
        string a = R(p[0]) + "," + R(p[1]) + "," + R(p[2]), b = R(p[3]) + "," + R(p[4]) + "," + R(p[5]);
        return string.CompareOrdinal(a, b) < 0 ? a + "|" + b : b + "|" + a;
    }

    /// <summary>Ayna ailesi: 4 aynadan (x, y işaretleri) en küçük anahtar.</summary>
    static string Aile(double[] p)
    {
        var l = new List<string>();
        foreach (double fx in new[] { 1.0, -1.0 })
            foreach (double fy in new[] { 1.0, -1.0 })
                l.Add(Anahtar(new[] { fx * p[0], fy * p[1], p[2], fx * p[3], fy * p[4], p[5] }));
        l.Sort(StringComparer.Ordinal);
        return l[0];
    }

    static double[] Koord(FinalMember m) { return new[] { m.start_x, m.start_y, m.start_z, m.end_x, m.end_y, m.end_z }; }

    public static int Main(string[] a)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        string yol = a[0]; bool ayrinti = a.Length > 1;
        var asm = typeof(KuleOkuyucu).Assembly; var t = asm.GetType("OutlineDrawing.dataprocessing");
        var o = Activator.CreateInstance(t, true);
        t.GetMethod("datapreprocessing", H).Invoke(o, new object[] { yol });
        var uyeler = KuleOkuyucu.Kopyala(o);
        int hata = 0;

        // ---- 1. okuma karşılaştırması
        Dictionary<string, int> diger; int eksikDugum;
        var tow = TowOku(yol, out diger, out eksikDugum);
        var odGeo = new HashSet<string>(uyeler.Values.Select(m => Anahtar(Koord(m))));
        var towGeo = new HashSet<string>(tow.Select(u => Anahtar(u.P)));
        var dusen = tow.Where(u => !odGeo.Contains(Anahtar(u.P))).ToList();
        int fazla = odGeo.Count(g => !towGeo.Contains(g));
        Console.WriteLine("  1. Okuma: .tow {0} eleman (benzersiz geometri {1}), OutlineDrawing {2} (benzersiz {3}); OutlineDrawing'de olmayan {4}, fazla {5}, düğümü bulunamayan {6}",
            tow.Count, towGeo.Count, uyeler.Count, odGeo.Count, dusen.Count, fazla, eksikDugum);
        foreach (var u in dusen.Take(ayrinti ? 1000 : 5)) Console.WriteLine("       OutlineDrawing'de yok: {0} ({1})", u.Ad, u.Grup);
        if (diger.Count > 0) Console.WriteLine("     UYARI: açı elemanı olmayan elemanlar (OutlineDrawing çizmez): " + string.Join(", ", diger.Select(kv => kv.Key + " " + kv.Value)));
        hata += dusen.Count + eksikDugum;

        // ---- 2. otomatik kesitler + gerçek SectionDetection
        var gorunen = KuleOkuyucu.Gorunenler(o);
        AutoSectionReport rapor;
        var secs = AutoSectionDetector.DetectHidden(uyeler, gorunen, new AutoSectionOptions(), out rapor);
        Type rowT = asm.GetType("OutlineDrawing.SectionRowData"), ptT = asm.GetType("OutlineDrawing.SectionPoint3D");
        var grid = (IDictionary)Activator.CreateInstance(typeof(Dictionary<,>).MakeGenericType(typeof(string), rowT));
        int ri = 0;
        foreach (var c in secs)
        {
            string ad = HarfAdi(ri);
            object row = Activator.CreateInstance(rowT);
            rowT.GetProperty("SectionName").SetValue(row, ad, null);
            rowT.GetProperty("RowIndex").SetValue(row, ri++, null);
            var pts = (IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(ptT));
            if (c.Kind == AutoSectionKind.Height) rowT.GetProperty("Height").SetValue(row, (double?)c.Height.Value, null);
            else foreach (var p in c.Points) pts.Add(Activator.CreateInstance(ptT, p.X, p.Y, p.Z));
            rowT.GetProperty("Points").SetValue(row, pts, null);
            grid[ad] = row;
        }
        t.GetMethod("SectionDetection", H).Invoke(o, new object[] { grid, yol, 15 });
        var kesitte = new HashSet<string>();
        foreach (DictionaryEntry de in (IDictionary)t.GetProperty("section").GetValue(o, null))
        {
            string k = Convert.ToString(de.Key, CultureInfo.InvariantCulture);
            int i = k.IndexOf(KesitAnahtari.Ayrac);
            kesitte.Add(i >= 0 ? k.Substring(i + 1) : k);
        }

        // ---- 3. görünüşler ve iç eleman izdüşümleri
        var on = KuleOkuyucu.Yuz(o, "frontFace"); var yan = KuleOkuyucu.Yuz(o, "sideFace");
        var onda = new HashSet<string>(on.Select(c => c.Anahtar)); var yanda = new HashSet<string>(yan.Select(c => c.Anahtar));
        var ic = new HashSet<string>(IcElemanPlani.Hesapla(uyeler, on, false).Concat(IcElemanPlani.Hesapla(uyeler, yan, true)).Select(c => c.Anahtar));

        // ---- aile bazında sonuç
        var aileler = uyeler.GroupBy(kv => Aile(Koord(kv.Value))).ToList();
        int etiketli = 0, yalnizIc = 0, yok = 0, kesitAile = 0, yuzAile = 0;
        var yokListe = new List<string>(); var icListe = new List<string>();
        foreach (var g in aileler)
        {
            bool yuzde = g.Any(kv => onda.Contains(kv.Key) || yanda.Contains(kv.Key));
            bool kesitDe = g.Any(kv => kesitte.Contains(kv.Key));
            if (yuzde) yuzAile++; else if (kesitDe) kesitAile++;
            if (yuzde || kesitDe) { etiketli++; continue; }
            string ornek = g.First().Key + " (" + g.First().Value.group_label + ", z " + Math.Round(g.First().Value.start_z) + "-" + Math.Round(g.First().Value.end_z) + ")";
            if (g.Any(kv => ic.Contains(kv.Key))) { yalnizIc++; icListe.Add(ornek); }
            else { yok++; yokListe.Add(ornek); }
        }
        Console.WriteLine("  2. Kesit: {0} otomatik kesit; görünmeyen {1} elemandan {2}'i kesitte (kapsanmayan {3})",
            secs.Count, rapor.HiddenMembers, rapor.CoveredHidden, rapor.UncoveredKeys.Count);
        Console.WriteLine("  3. Eleman aileleri (ayna kopyalar bir aile): {0}; görünüşte {1}, kesitte {2} -> etiketli çizilen {3}; yalnız iç eleman çizgisi {4}; HİÇ ÇİZİLMEYEN {5}",
            aileler.Count, yuzAile, kesitAile, etiketli, yalnizIc, yok);
        foreach (var s in icListe.Take(ayrinti ? 1000 : 8)) Console.WriteLine("       yalnız iç eleman çizgisi: " + s);
        foreach (var s in yokListe.Take(ayrinti ? 1000 : 8)) Console.WriteLine("       HİÇ ÇİZİLMİYOR: " + s);
        hata += yok;
        Console.WriteLine(hata == 0 ? "  SONUC: atlanan eleman yok" : "  SONUC: " + hata + " sorun");
        return hata == 0 ? 0 : 1;
    }

    static string HarfAdi(int i)
    {
        string s = "";
        i++;
        while (i > 0) { int r = (i - 1) % 26; s = (char)('A' + r) + s; i = (i - 1) / 26; }
        return s;
    }
}
