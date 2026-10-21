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
    static string HarfAdi(int i)
    {
        string s = "";
        i++;
        while (i > 0) { int r = (i - 1) % 26; s = (char)('A' + r) + s; i = (i - 1) / 26; }
        return s;
    }

    /// <summary>Kesitin paftadaki yeri: çizgiler, ölçü yazıları ve başlık t, (du, dv) kaymayla; cizgi / yazi verilirse önizlemeye
    /// yazar. Kapladığı kutu döner.</summary>
    static double[] Gercek(int i, List<double[]> segs, double kay, double[] c, double[] t, double du, double dv,
                           Action<string, double[]> cizgi = null, Action<double, double, double, string> yazi = null, string baslik = null)
    {
        Func<double, double> X = x => x + du;
        Func<double, double> Yv = y => y + dv;
        double[] r = null;
        Action<double[]> ekle = q => r = r == null ? (double[])q.Clone() : new[] { Math.Min(r[0], q[0]), Math.Min(r[1], q[1]), Math.Max(r[2], q[2]), Math.Max(r[3], q[3]) };
        foreach (var q in segs)
        {
            var l = new[] { X(q[0] + kay), Yv(q[1]), X(q[2] + kay), Yv(q[3]) };
            ekle(new[] { Math.Min(l[0], l[2]), Math.Min(l[1], l[3]), Math.Max(l[0], l[2]), Math.Max(l[1], l[3]) });
            if (cizgi != null) cizgi("L", l);
        }
        foreach (var o in OlcuPlani.Hesapla(segs))
        {
            var b = Kutu(o);
            double mx = X((b[0] + b[2]) / 2 + kay), my = Yv((b[1] + b[3]) / 2), hw = (b[2] - b[0]) / 2, hh = (b[3] - b[1]) / 2;
            var d = new[] { mx - hw, my - hh, mx + hw, my + hh };
            ekle(d);
            if (cizgi != null) cizgi("D", d);
        }
        double tu = (t[0] + t[2]) / 2 + du, tv0 = t[1] + dv, tw = (t[2] - t[0]) / 2;
        ekle(new[] { tu - tw, tv0, tu + tw, t[3] + dv });
        if (yazi != null) yazi(tu, tv0, 400, baslik);
        return r;
    }

    // ---- eleman isimleri (EtiketCozucu): OutlineDrawing'in yerleşimi benzetilir (her grubun ismi bir elemanında: +X yarıda en
    //      sağdaki; görünüşte eleman doğrultusunda, 0,15 boy yukarıda, yazı ortası / alt; kesitte yatay, elemanın ortasından sola
    //      hizalı), sonra çakışmalar giderilir. Yazı genişliği karakter başına 0,75 boy, isim boyu 150.
    static double EtiketBoyu = 150, KarakterOrani = 0.75;

    static int EtiketTesti(KuleVerisi k, List<KesitTanimi> tanimlar, List<string> rapor)
    {
        double h = EtiketBoyu, pay = 0.05 * h;
        int kalanToplam = 0;
        string onizleme = Environment.GetEnvironmentVariable("OTO_ETIKET_ONIZLEME");
        System.IO.StreamWriter w = string.IsNullOrEmpty(onizleme) ? null : new System.IO.StreamWriter(onizleme);
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        Action<string, double[]> yaz = (tur, q) => { if (w != null) w.WriteLine(tur + " " + string.Join(" ", q.Select(x => x.ToString("0.#", ci)))); };
        // görünüşler
        foreach (var g in new[] { new { ad = "on", y = k.OnYuz, yan = false, du = 0.0 }, new { ad = "yan", y = k.YanYuz, yan = true, du = 60000.0 } })
        {
            var cizgiler = g.y.Concat(IcElemanPlani.Hesapla(k.Uyeler, g.y, g.yan)).Select(c => new[] { c.P[0] + g.du, c.P[2], c.P[3] + g.du, c.P[5] }).ToList();
            var etiketler = new List<EtiketCozucu.Etiket>();
            bool esli = Environment.GetEnvironmentVariable("OTO_ETIKET_ES") != "0";
            foreach (var grup in g.y.Where(c => !string.IsNullOrEmpty(c.Grup)).GroupBy(c => c.Grup.Trim(), StringComparer.OrdinalIgnoreCase))
            {
                var arti = grup.Where(c => c.P[0] >= 0 && c.P[3] >= 0).ToList();
                YuzCizgisi m = (arti.Count > 0 ? arti : grup.ToList()).OrderByDescending(c => (c.P[0] + c.P[3]) / 2).First();
                double[] s = { m.P[0] + g.du, m.P[2], m.P[3] + g.du, m.P[5] };
                int dayanak = g.y.IndexOf(m);
                double dx = s[2] - s[0], dz = s[3] - s[1], L = Math.Sqrt(dx * dx + dz * dz);
                if (L < 1e-6 || dayanak < 0) continue;
                double nx = -dz / L, nz = dx / L, a = Math.Atan2(dz, dx);
                if (a > Math.PI / 2 || a < -Math.PI / 2) a += Math.PI;
                double tu = Math.Cos(a), tv = Math.Sin(a), uu = -tv, uv = tu;
                double qu = (s[0] + s[2]) / 2 + nx * 0.15 * h, qv = (s[1] + s[3]) / 2 + nz * 0.15 * h;
                double ww = grup.Key.Length * KarakterOrani * h;
                var et = new EtiketCozucu.Etiket();
                // aynı grubun öbür elemanları (EtiketDuzeni.GrupEsleri gibi), yakından uzağa
                double mu = (s[0] + s[2]) / 2, mv = (s[1] + s[3]) / 2;
                var hostlar = new List<int> { dayanak };
                if (esli) hostlar.AddRange(grup.Select(c => g.y.IndexOf(c)).Where(j => j != dayanak)
                    .OrderBy(j => Math.Abs((cizgiler[j][0] + cizgiler[j][2]) / 2 - mu) + Math.Abs((cizgiler[j][1] + cizgiler[j][3]) / 2 - mv)).Take(12));
                EtiketCozucu.AdaylarEkle(et, hostlar.Select(j => cizgiler[j]).ToList(), hostlar, qu, qv, tu, tv, ww, h, qu + uu * h / 2, qv + uv * h / 2, true);
                etiketler.Add(et);
            }
            kalanToplam += EtiketSay(g.ad, etiketler, cizgiler, pay, rapor, yaz);
        }
        // kesitler (yatay ya da düzlem kesiti; OutlineDrawing etiketi): tam travers ve döndürülen kesitler hariç
        int kesitOnce = 0, kesitSonra = 0, kesitIsim = 0, kesitTasinan = 0;
        for (int i = 0; i < tanimlar.Count; i++)
        {
            if (TraversKesiti.Hesapla(k.Uyeler, tanimlar[i]) != null || KesitYonu.Aci(tanimlar[i]) != 0) continue;
            List<string> anahtarlar;
            var segs = KesitSecici.Sec(k.Uyeler, tanimlar[i], out anahtarlar);
            var etiketler = new List<EtiketCozucu.Etiket>();
            foreach (var grup in Enumerable.Range(0, segs.Count).GroupBy(j => (k.Uyeler[anahtarlar[j]].group_label ?? "").Trim(), StringComparer.OrdinalIgnoreCase))
            {
                if (grup.Key.Length == 0) continue;
                Func<int, double> midX = j => (k.Uyeler[anahtarlar[j]].start_x + k.Uyeler[anahtarlar[j]].end_x) / 2;
                var arti = grup.Where(j => k.Uyeler[anahtarlar[j]].start_x >= 0 && k.Uyeler[anahtarlar[j]].end_x >= 0).ToList();
                int m = (arti.Count > 0 ? arti : grup.ToList()).OrderByDescending(midX).First();
                double[] s = segs[m];
                double qu = (s[0] + s[2]) / 2, qv = (s[1] + s[3]) / 2, ww = grup.Key.Length * KarakterOrani * h;
                var et = new EtiketCozucu.Etiket();
                EtiketCozucu.AdaylarEkle(et, new List<double[]> { s }, new List<int> { m }, qu, qv, 1, 0, ww, h, qu + ww / 2, qv, false);
                etiketler.Add(et);
            }
            int once = Cakisan(etiketler.Select(e => e.Adaylar[0].K).ToList(), pay);
            int tasinan = EtiketCozucu.Coz(etiketler, segs, pay);
            int sonra = Cakisan(etiketler.Select(e => e.Adaylar[e.Secim].K).ToList(), pay);
            kesitOnce += once; kesitSonra += sonra; kesitIsim += etiketler.Count; kesitTasinan += tasinan;
            if (sonra > 0) rapor.Add("etiket: kesit " + HarfAdi(i) + " " + sonra + " cakisma kaldi");
        }
        Console.WriteLine("Etiket kesitler: {0} isim, cakisma once {1}, sonra {2} ({3} isim tasindi)", kesitIsim, kesitOnce, kesitSonra, kesitTasinan);
        kalanToplam += kesitSonra;
        if (w != null) w.Dispose();
        return kalanToplam;
    }

    static int Cakisan(List<double[]> kutular, double pay)
    {
        int n = 0;
        for (int i = 0; i < kutular.Count; i++)
            for (int j = i + 1; j < kutular.Count; j++)
                if (EtiketCozucu.Biner(kutular[i], kutular[j], pay)) n++;
        return n;
    }

    static int EtiketSay(string ad, List<EtiketCozucu.Etiket> etiketler, List<double[]> cizgiler, double pay, List<string> rapor, Action<string, double[]> yaz)
    {
        int once = Cakisan(etiketler.Select(e => e.Adaylar[0].K).ToList(), pay);
        int ustunde = etiketler.Count(e => e.Zorla);
        int tasinan = EtiketCozucu.Coz(etiketler, cizgiler, pay);
        int sonra = Cakisan(etiketler.Select(e => e.Adaylar[e.Secim].K).ToList(), pay);
        int ustundeSonra = etiketler.Count(e => EtiketCozucu.Keser(e.Adaylar[e.Secim].K, cizgiler[e.Dayanak], pay));
        Console.WriteLine("Etiket {0}: {1} isim, cakisma once {2} (kendi cizgisi ustunde {3}), sonra {4} (ustunde {5}); {6} isim tasindi",
            ad, etiketler.Count, once, ustunde, sonra, ustundeSonra, tasinan);
        if (sonra > 0) rapor.Add("etiket: " + ad + " gorunuste " + sonra + " cakisma kaldi");
        if (ad == "on")
        {
            foreach (var c in cizgiler) yaz("L", c);
            foreach (var e in etiketler) { yaz("A", e.Adaylar[0].K); yaz("B", e.Adaylar[e.Secim].K); }
        }
        return sonra + ustundeSonra;
    }

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
        // yan görünüş çizilmiyorsa (OTO_YAN=0) pafta yalnız ön görünüşle
        bool yanVar = Environment.GetEnvironmentVariable("OTO_YAN") != "0";
        var uzatma = UzatmaPlani.Hesapla(AileParcalari.Oku(TowYama.SatirlarHam(a[0])));
        var liste = MalzemeListesi.Oku(a[0]);
        if (Environment.GetEnvironmentVariable("OTO_LISTE") == "0") liste.Clear();   // listesiz pafta benzetimi
        // pafta boşluğuna ek kesitler (Oto Kesit, "Paftayı doldur"; OTO_EK=0 ise yok)
        if (Environment.GetEnvironmentVariable("OTO_EK") != "0")
        {
            var adaylar = OtoKesitBulucu.EkAdaylar(k.Uyeler, k.GorunenAnahtarlar, secs);
            double s0;
            var ek = PaftaTahmini.Doldur(k, secs, adaylar, uzatma.Select(c => c.Cizgiler).ToList(), uzatma.Select(c => c.Baslik).ToList(), yanVar, liste.Count > 0, out s0);
            Console.WriteLine("Ek kesit: {0} aday, {1} eklendi (tahmini olcek 1:{2:0})", adaylar.Count, ek.Count, s0);
            foreach (var c in adaylar) Console.WriteLine("  {0} {1}", ek.Contains(c) ? "+" : "-", c.Description);
            secs = OtoKesitBulucu.Sirala(secs.Concat(ek));
        }
        var rapor = new List<string>();
        int toplam = 0, sig = 0, cak = 0, ki = 0;
        var kesitKutulari = new List<double[]>();
        var cizimKutulari = new List<double[]>(); var baslikKutulari = new List<double[]>();   // başlık hariç çizim / başlık (büyütme)
        // çizimdeki kesitler: OutlineDrawing'in (tam travers kesitleri eklentinin açılmış çizimiyle) + aile dosyasında
        // eklentinin çizdiği uzatma görünüşleri ve K / L kesitleri
        var tanimlar = secs.Select(c => OtoKesitBulucu.Tanim(c)).ToList();
        var cizimler = new List<List<double[]>>();
        for (int i = 0; i < tanimlar.Count; i++)
        {
            var segs = KesitSecici.CizimSegmentleri(k.Uyeler, tanimlar[i]);
            // yatay kesitlerden uzak küçük kopuk parçalar kaldırılır (KopukParcaTemizle)
            var at = KopukParca.Atilacaklar(k.Uyeler, tanimlar[i], tanimlar.Where((x, j) => j != i));
            if (at.Count > 0) segs = segs.Where((q, j) => !at.Contains(j)).ToList();
            // plana yakın eğik kesit (tam travers değilse) KesitYonu ile 90° döndürülür; ölçüler döndürülmüş hale konur
            int aci = TraversKesiti.Hesapla(k.Uyeler, tanimlar[i]) == null ? KesitYonu.Aci(tanimlar[i]) : 0;
            if (aci != 0 && segs.Count > 0)
            {
                double cu = (segs.Min(q => Math.Min(q[0], q[2])) + segs.Max(q => Math.Max(q[0], q[2]))) / 2;
                double cv = (segs.Min(q => Math.Min(q[1], q[3])) + segs.Max(q => Math.Max(q[1], q[3]))) / 2;
                double sg = aci > 0 ? 1 : -1;   // +90: (u, v) -> (-v, u)
                segs = segs.Select(q => new[] { cu - sg * (q[1] - cv), cv + sg * (q[0] - cu), cu - sg * (q[3] - cv), cv + sg * (q[2] - cu) }).ToList();
            }
            cizimler.Add(segs);
        }
        var basliklar0 = Enumerable.Range(0, tanimlar.Count).Select(i => "SECTION " + HarfAdi(i)).ToList();
        cizimler.AddRange(uzatma.Select(c => c.Cizgiler));
        basliklar0.AddRange(uzatma.Select(c => c.Baslik));
        var kutuCizim = new List<int>(); var kaylar = new List<double>();
        if (uzatma.Count > 0) Console.WriteLine("Aile: {0} uzatma cizimi ({1})", uzatma.Count, string.Join(", ", uzatma.Select(c => c.Baslik)));
        for (int ci = 0; ci < cizimler.Count; ci++)
        {
            var segs = cizimler[ci];
            var l = OlcuPlani.Hesapla(segs);
            toplam += l.Count;
            var s = Say(l, "kesit " + (++ki), rapor); sig += s[0]; cak += s[1];
            // dizim testi için kesitin kapladığı alan: çizgiler + ölçü yazıları + altında 400'lük başlık (OutlineDrawing)
            if (segs.Count == 0) continue;
            double u0 = segs.Min(q => Math.Min(q[0], q[2])), u1 = segs.Max(q => Math.Max(q[0], q[2]));
            double v0 = segs.Min(q => Math.Min(q[1], q[3])), v1 = segs.Max(q => Math.Max(q[1], q[3]));
            foreach (var o in l) { var b = Kutu(o); u0 = Math.Min(u0, b[0]); v0 = Math.Min(v0, b[1]); u1 = Math.Max(u1, b[2]); v1 = Math.Max(v1, b[3]); }
            // OutlineDrawing'in tek sıra düzeni: kesit i, kesit aralığı kadar sağda
            double kay = (kesitKutulari.Count) * k.KesitAraligi;
            // başlık (400): çizimin ortasında, çizimden 600 aşağıda; kutu başlığın altında 400 pay
            double baslikW = basliklar0[ci].Length * 0.8 * 400, um = (segs.Min(q => Math.Min(q[0], q[2])) + segs.Max(q => Math.Max(q[0], q[2]))) / 2;
            cizimKutulari.Add(new[] { u0 + kay, v0, u1 + kay, v1 });
            baslikKutulari.Add(new[] { um - baslikW / 2 + kay, v0 - 1000, um + baslikW / 2 + kay, v0 - 600 });
            u0 = Math.Min(u0, um - baslikW / 2); u1 = Math.Max(u1, um + baslikW / 2); v0 -= 1400;
            kesitKutulari.Add(new[] { u0 + kay, v0, u1 + kay, v1 });
            kutuCizim.Add(ci); kaylar.Add(kay);
        }
        // ---- gruplama (KesitDizici.Grupla): ÇALIŞTIR sonundaki yerleşimde (OutlineDrawing kesitleri kesit aralığıyla, aile
        //      uzatma çizimleri en sağda sık) her çizgi ve ölçü kendi kesit başlığına bağlanmalı. Eski tek pencere (en sık iki
        //      başlığın aralığı) geniş kesitlerin uçlarını bırakıyordu; şimdiki pencere başlık başına.
        {
            var merkez = new List<double>(); var nesneU = new List<List<double>>();
            double ox = 0;
            for (int i = 0; i < cizimler.Count; i++)
            {
                var segs = cizimler[i];
                var us = new List<double>();
                if (segs.Count == 0) { merkez.Add(double.NaN); nesneU.Add(us); continue; }
                double u0 = segs.Min(q => Math.Min(q[0], q[2])), u1 = segs.Max(q => Math.Max(q[0], q[2]));
                double kay;
                if (i < secs.Count) { kay = i * k.KesitAraligi; ox = Math.Max(ox, u1 + kay); }
                else { if (i == secs.Count) ox += 10000; kay = ox - u0; ox += (u1 - u0) + 6000; }   // UzatmaCizimleri ile aynı
                foreach (var q in segs) us.Add((q[0] + q[2]) / 2 + kay);
                foreach (var o in OlcuPlani.Hesapla(segs)) { var bk = Kutu(o); us.Add((bk[0] + bk[2]) / 2 + kay); }
                merkez.Add((u0 + u1) / 2 + kay);
                nesneU.Add(us);
            }
            var basliklar = Enumerable.Range(0, merkez.Count).Where(i => !double.IsNaN(merkez[i])).OrderBy(i => merkez[i]).ToList();
            var pencere = new Dictionary<int, double>();
            double enKucuk = double.MaxValue;
            for (int j = 0; j < basliklar.Count; j++)
            {
                double sol = j > 0 ? merkez[basliklar[j]] - merkez[basliklar[j - 1]] : 0, sag = j + 1 < basliklar.Count ? merkez[basliklar[j + 1]] - merkez[basliklar[j]] : 0;
                double w = Math.Max(sol, sag) / 2;
                pencere[basliklar[j]] = w > 500 ? w : 15000;
                if (j > 0 && sol > 1000) enKucuk = Math.Min(enKucuk, sol);
            }
            int yeniKopuk = 0, eskiKopuk = 0;
            foreach (int i in basliklar)
                foreach (double u in nesneU[i])
                {
                    int en = basliklar.OrderBy(j => Math.Abs(merkez[j] - u)).First();
                    if (en != i || Math.Abs(merkez[en] - u) > pencere[en])
                    {
                        yeniKopuk++;
                        rapor.Add(string.Format(System.Globalization.CultureInfo.InvariantCulture,
                            "gruplama: cizim {0} (merkez {1:0}) nesnesi u={2:0} -> en yakin cizim {3} (merkez {4:0}, pencere {5:0})", i + 1, merkez[i], u, en + 1, merkez[en], pencere[en]));
                    }
                    if (en != i || Math.Abs(merkez[en] - u) > enKucuk / 2) eskiKopuk++;
                }
            Console.WriteLine("Gruplama: {0} cizim, kendi basligina baglanmayan nesne {1} (eski tek pencereyle {2})", basliklar.Count, yeniKopuk, eskiKopuk);
            if (yeniKopuk > 0) rapor.Add("gruplama: " + yeniKopuk + " nesne kendi kesitine baglanmiyor");
            cak += yeniKopuk;
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

        // ---- pafta: görünüşler + kesitler şablonun iç çerçevesine, malzeme listesi ve antetle çakışmadan
        Func<List<double[]>, double[]> kutuL = L => new[] { L.Min(q => Math.Min(q[0], q[2])) - 300, L.Min(q => Math.Min(q[1], q[3])) - 1600, L.Max(q => Math.Max(q[0], q[2])) + 2500, L.Max(q => Math.Max(q[1], q[3])) + 300 };
        Func<List<YuzCizgisi>, bool, int, List<double[]>> yuz = (F, yanMi, ofs) => F.Concat(IcElemanPlani.Hesapla(k.Uyeler, F, yanMi)).Select(c => new[] { c.P[0] + ofs, c.P[2], c.P[3] + ofs, c.P[5] }).ToList();
        double[] onK = kutuL(yuz(k.OnYuz, false, 0)), yanK = yanVar ? kutuL(yuz(k.YanYuz, true, k.GorunusAraligi)) : null;
        // gerçek pafta ön görünüş kutusuna kesit işaretlerini de katar (KesitIsaretleri, SECTION_MARK)
        var tan = new Dictionary<string, KesitTanimi>();
        for (int i = 0; i < tanimlar.Count; i++) tan[HarfAdi(i)] = tanimlar[i];
        var iki = new HashSet<string>(tan.Where(kv => TraversKesiti.Hesapla(k.Uyeler, kv.Value) != null).Select(kv => kv.Key));
        var isaretler = KesitIsaretPlani.Hesapla(k.OnYuz, tan, iki, k.Uyeler);
        foreach (var m in isaretler) onK = new[] { Math.Min(onK[0], m.Kutu[0]), Math.Min(onK[1], m.Kutu[1]), Math.Max(onK[2], m.Kutu[2]), Math.Max(onK[3], m.Kutu[3]) };
        var pf = PaftaPlani.Hesapla(onK, yanK, kesitKutulari, liste.Count > 0, 841, 594.65);
        int paftaSorun = 0;
        if (!pf.Sigdi) { paftaSorun++; rapor.Add("pafta: sigmadi"); }
        else
        {
            Func<double[], double, double, double[]> kagit = (b, du, dv) => new[] { (b[0] + du - pf.OrjinU) / pf.Olcek, (b[1] + dv - pf.OrjinV) / pf.Olcek, (b[2] + du - pf.OrjinU) / pf.Olcek, (b[3] + dv - pf.OrjinV) / pf.Olcek };
            var parca = new List<double[]> { kagit(onK, 0, 0) };
            if (yanK != null) parca.Add(kagit(yanK, pf.YanDu, 0));
            for (int i = 0; i < kesitKutulari.Count; i++) parca.Add(kagit(kesitKutulari[i], pf.KesitKayma[i][0], pf.KesitKayma[i][1]));
            double G = 841, Y = 594.65;
            double[] ic = { PaftaPlani.Ic[0] * G, PaftaPlani.Ic[1] * Y, PaftaPlani.Ic[2] * G, PaftaPlani.Ic[3] * Y };
            var yasak = new List<double[]> { new[] { PaftaPlani.Antet[0] * G, PaftaPlani.Antet[1] * Y, PaftaPlani.Antet[2] * G, PaftaPlani.Antet[3] * Y } };
            if (liste.Count > 0) yasak.Add(new[] { PaftaPlani.Liste[0] * G, PaftaPlani.Liste[1] * Y, PaftaPlani.Liste[2] * G, PaftaPlani.Liste[3] * Y });
            Func<double[], double[], bool> biner = (x, y) => x[0] < y[2] - 0.01 && y[0] < x[2] - 0.01 && x[1] < y[3] - 0.01 && y[1] < x[3] - 0.01;
            for (int i = 0; i < parca.Count; i++)
            {
                var b = parca[i];
                if (b[0] < ic[0] - 0.01 || b[1] < ic[1] - 0.01 || b[2] > ic[2] + 0.01 || b[3] > ic[3] + 0.01) { paftaSorun++; rapor.Add("pafta: parca " + i + " cerceve disinda"); }
                foreach (var y in yasak) if (biner(b, y)) { paftaSorun++; rapor.Add("pafta: parca " + i + " liste/antet ustunde"); }
                for (int j = i + 1; j < parca.Count; j++) if (biner(b, parca[j])) { paftaSorun++; rapor.Add("pafta: parca " + i + " / " + j + " biniyor"); }
            }
            Console.WriteLine("Pafta: olcek 1:{0:0}, {1} parca, malzeme listesi {2} satir, sorun {3}", pf.Olcek, parca.Count, liste.Count, paftaSorun);
            // ---- önizleme: OTO_ONIZLEME=dosya ise paftanın kâğıt koordinatlarında (mm) çizgileri
            string onizleme = Environment.GetEnvironmentVariable("OTO_ONIZLEME");
            if (!string.IsNullOrEmpty(onizleme))
            {
                var ci2 = System.Globalization.CultureInfo.InvariantCulture;
                using (var w = new System.IO.StreamWriter(onizleme))
                {
                    Action<string, double[]> yaz = (tur, q) => w.WriteLine(tur + " " + string.Join(" ", q.Select(x => x.ToString("0.###", ci2))));
                    Func<double, double, double[]> P = (u, v) => new[] { (u - pf.OrjinU) / pf.Olcek, (v - pf.OrjinV) / pf.Olcek };
                    Action<string, double[], double, double> cizgi = (tur, q, du, dv) => { var a1 = P(q[0] + du, q[1] + dv); var b1 = P(q[2] + du, q[3] + dv); yaz(tur, new[] { a1[0], a1[1], b1[0], b1[1] }); };
                    yaz("R", ic); foreach (var y in yasak) yaz("R", y);
                    foreach (var c in k.OnYuz) cizgi("L", new[] { c.P[0], c.P[2], c.P[3], c.P[5] }, 0, 0);
                    foreach (var c in IcElemanPlani.Hesapla(k.Uyeler, k.OnYuz, false)) cizgi("H", new[] { c.P[0], c.P[2], c.P[3], c.P[5] }, 0, 0);
                    if (yanK != null)
                    {
                        foreach (var c in k.YanYuz) cizgi("L", new[] { c.P[0], c.P[2], c.P[3], c.P[5] }, k.GorunusAraligi + pf.YanDu, 0);
                        foreach (var c in IcElemanPlani.Hesapla(k.Uyeler, k.YanYuz, true)) cizgi("H", new[] { c.P[0], c.P[2], c.P[3], c.P[5] }, k.GorunusAraligi + pf.YanDu, 0);
                    }
                    foreach (var m in isaretler)
                    {
                        foreach (var c in m.Cizgiler) cizgi("M", c, 0, 0);
                        var t = P(m.YaziU, m.YaziV); w.WriteLine("T {0} {1} {2} {3}", t[0].ToString("0.###", ci2), t[1].ToString("0.###", ci2), (m.YaziBoyu / pf.Olcek).ToString("0.###", ci2), m.Ad);
                    }
                    double onOrta = (onK[0] + onK[2]) / 2;
                    var to = P(onOrta, onK[1] + 600); w.WriteLine("T {0} {1} {2} TRANSVERSE FACE", to[0].ToString("0.###", ci2), to[1].ToString("0.###", ci2), (400 / pf.Olcek).ToString("0.###", ci2));
                    if (yanK != null)
                    {
                        double yanOrta = (yanK[0] + yanK[2]) / 2 + pf.YanDu;
                        var ty = P(yanOrta, yanK[1] + 600); w.WriteLine("T {0} {1} {2} LONGITUDINAL FACE", ty[0].ToString("0.###", ci2), ty[1].ToString("0.###", ci2), (400 / pf.Olcek).ToString("0.###", ci2));
                    }
                    for (int i = 0; i < kesitKutulari.Count; i++)
                    {
                        Gercek(i, cizimler[kutuCizim[i]], kaylar[i], cizimKutulari[i], baslikKutulari[i],
                               pf.KesitKayma[i][0], pf.KesitKayma[i][1], (tur, q) => cizgi(tur, q, 0, 0),
                               (u, v, h, m) => { var t = P(u, v); w.WriteLine("T {0} {1} {2} {3}", t[0].ToString("0.###", ci2), t[1].ToString("0.###", ci2), (h / pf.Olcek).ToString("0.###", ci2), m); },
                               basliklar0[kutuCizim[i]]);
                    }
                    for (int i = 0; i < liste.Count && pf.ListeAlani != null; i++) { }
                    if (pf.ListeAlani != null) yaz("Q", pf.ListeAlani);
                    w.WriteLine("S 1:{0:0}", pf.Olcek);
                }
            }
        }
        cak += paftaSorun;
        EtiketTesti(k, tanimlar, rapor);
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
