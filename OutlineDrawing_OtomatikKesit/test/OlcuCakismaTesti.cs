// Ölçü yazısı çakışma testi (AutoCAD'siz): otomatik kesitlerin ve ön / yan görünüşün ölçüleri için yazı kutuları
// tahmin edilir (AutoCAD "best fit": yazı iki uzatma çizgisi arasına sığıyorsa ortada, ölçü çizgisinin üstünde; elle
// konum verilmişse orada). Sığmayıp elle konumu olmayan yazı (AutoCAD dışarı atar, komşusuna biner) ve birbirine
// binen yazı sayılır.
//   bash test/olcu_testi.sh eklenti/bin/tek/OutlineDrawing.dll kule.tow
using System; using System.Collections.Generic; using System.Linq; using System.Reflection; using OtomatikKesit;

public static class OlcuCakismaTesti
{
    const BindingFlags H = BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;

    // { u0, v0, u1, v1 }
    static double[] Kutu(OlcuTanimi o)
    {
        double w = OlcuPlani.YaziGenisligi(o), h = o.Yazi, gap = 0.25 * o.Yazi;
        bool yatay = Math.Abs(o.Rotation) < 1e-9;
        if (o.YaziU.HasValue && o.YaziV.HasValue)
            return yatay ? new[] { o.YaziU.Value - w / 2, o.YaziV.Value - h / 2, o.YaziU.Value + w / 2, o.YaziV.Value + h / 2 }
                         : new[] { o.YaziU.Value - h / 2, o.YaziV.Value - w / 2, o.YaziU.Value + h / 2, o.YaziV.Value + w / 2 };
        if (yatay)
        {
            double um = (o.U1 + o.U2) / 2, vb = o.VD + gap;
            return new[] { um - w / 2, vb, um + w / 2, vb + h };
        }
        double vm = (o.V1 + o.V2) / 2, ur = o.UD - gap;
        return new[] { ur - h, vm - w / 2, ur, vm + w / 2 };
    }

    public static bool Sigar(OlcuTanimi o) { return o.Deger >= OlcuPlani.YaziGenisligi(o) + 0.5 * o.Yazi + 0.2 * o.Yazi; }

    static int Elle;

    static int[] Say(List<OlcuTanimi> l, string ad, List<string> rapor)
    {
        int sigmayan = 0, cakisan = 0;
        Elle += l.Count(o => o.YaziU.HasValue && o.YaziV.HasValue);
        var kutular = l.Select(Kutu).ToList();
        for (int i = 0; i < l.Count; i++)
            if (!Sigar(l[i]) && !(l[i].YaziU.HasValue && l[i].YaziV.HasValue)) { sigmayan++; rapor.Add(ad + ": sigmayan " + Math.Round(l[i].Deger)); }
        for (int i = 0; i < l.Count; i++)
            for (int j = i + 1; j < l.Count; j++)
            {
                double[] a = kutular[i], b = kutular[j];
                double pay = 0.05 * Math.Min(l[i].Yazi, l[j].Yazi);
                if (a[0] < b[2] - pay && b[0] < a[2] - pay && a[1] < b[3] - pay && b[1] < a[3] - pay)
                { cakisan++; rapor.Add(ad + ": cakisan " + Math.Round(l[i].Deger) + " / " + Math.Round(l[j].Deger)); }
            }
        return new[] { sigmayan, cakisan };
    }

    public static int Main(string[] a)
    {
        var asm = typeof(KuleOkuyucu).Assembly;
        KuleVerisi k = KuleOkuyucu.OkuDogrudan(asm, a[0]);
        AutoSectionReport r;
        var secs = OtoKesitBulucu.Bul(k.Uyeler, k.GorunenAnahtarlar, out r);
        var rapor = new List<string>();
        int toplam = 0, sig = 0, cak = 0, ki = 0;
        var kesitKutulari = new List<double[]>();
        foreach (var c in secs)
        {
            var t = new KesitTanimi();
            if (c.Kind == AutoSectionKind.Height) t.Yukseklik = c.Height.Value; else foreach (var p in c.Points) t.Noktalar.Add(new[] { p.X, p.Y, p.Z });
            List<string> anah;
            var segs = KesitSecici.Sec(k.Uyeler, t, out anah);
            var l = OlcuPlani.Hesapla(segs);
            toplam += l.Count;
            var s = Say(l, "kesit " + (++ki), rapor); sig += s[0]; cak += s[1];
            // dizim testi için kesitin kapladığı alan: çizgiler + ölçü yazıları + altında 400'lük başlık (OutlineDrawing)
            if (segs.Count == 0) continue;
            double u0 = segs.Min(q => Math.Min(q[0], q[2])), u1 = segs.Max(q => Math.Max(q[0], q[2]));
            double v0 = segs.Min(q => Math.Min(q[1], q[3])), v1 = segs.Max(q => Math.Max(q[1], q[3]));
            foreach (var o in l) { var b = Kutu(o); u0 = Math.Min(u0, b[0]); v0 = Math.Min(v0, b[1]); u1 = Math.Max(u1, b[2]); v1 = Math.Max(v1, b[3]); }
            double baslikW = ("SECTION " + ki).Length * 0.8 * 400, um = (u0 + u1) / 2;
            u0 = Math.Min(u0, um - baslikW / 2); u1 = Math.Max(u1, um + baslikW / 2); v0 -= 1400;
            // OutlineDrawing'in tek sıra düzeni: kesit i, kesit aralığı kadar sağda
            double kay = (kesitKutulari.Count) * k.KesitAraligi;
            kesitKutulari.Add(new[] { u0 + kay, v0, u1 + kay, v1 });
        }
        // ---- kesit dizimi: ızgarada kesitler birbirine binmiyor, sıra korunuyor
        int satir; var kayma = KesitDizimi.Yerlestir(kesitKutulari, out satir);
        var yeni = kesitKutulari.Select((b, i) => new[] { b[0] + kayma[i][0], b[1] + kayma[i][1], b[2] + kayma[i][0], b[3] + kayma[i][1] }).ToList();
        int binen = 0;
        for (int i = 0; i < yeni.Count; i++)
            for (int j = i + 1; j < yeni.Count; j++)
                if (yeni[i][0] < yeni[j][2] && yeni[j][0] < yeni[i][2] && yeni[i][1] < yeni[j][3] && yeni[j][1] < yeni[i][3]) { binen++; rapor.Add("dizim: kesit " + (i + 1) + " / " + (j + 1) + " biniyor"); }
        if (yeni.Count > 0)
        {
            Func<List<double[]>, double> gen = L => L.Max(b => b[2]) - L.Min(b => b[0]);
            Func<List<double[]>, double> yuk = L => L.Max(b => b[3]) - L.Min(b => b[1]);
            Console.WriteLine("Dizim: {0} kesit, {1} satir; blok {2:0.0} x {3:0.0} m (tek sira {4:0.0} x {5:0.0} m), binen {6}",
                yeni.Count, satir, gen(yeni) / 1000, yuk(yeni) / 1000, gen(kesitKutulari) / 1000, yuk(kesitKutulari) / 1000, binen);
        }
        cak += binen;
        foreach (var g in new[] { new { ad = "on", y = k.OnYuz, o = 0 }, new { ad = "yan", y = k.YanYuz, o = k.GorunusAraligi } })
        {
            var l = GorunusOlcuPlani.Hesapla(g.y, g.o);
            toplam += l.Count;
            var s = Say(l, "gorunus " + g.ad, rapor); sig += s[0]; cak += s[1];
        }
        foreach (string x in rapor.Take(30)) Console.WriteLine("  " + x);
        Console.WriteLine("Olcu: {0} olcu (yazisi disari / uste alinan {3}), sigmayan (elle konumsuz) {1}, cakisan yazi {2}", toplam, sig, cak, Elle);
        Console.WriteLine(sig + cak == 0 ? "SONUC: olcu yazilari temiz" : "SONUC: olcu yazilarinda sorun");
        return sig + cak;
    }
}
