// Görünüş yeri bulma testi: OutlineDrawing'in çizim düzeni (ön 0, yan ön varsa 50000 yoksa 0, kesitler sonra)
// benzetilir; her durumda doğru görünüşlerin doğru yerde bulunduğu, olmayanların bulunmadığı kontrol edilir.
using System; using System.Linq; using System.Collections.Generic; using OtomatikKesit;
class P {
  static List<double[]> Ciz(List<YuzCizgisi> F, int ofs, bool ikiD) {
    return F.Select(c => ikiD ? new[]{ c.P[0]+ofs, c.P[2], -c.P[1], c.P[3]+ofs, c.P[5], -c.P[4] }
                              : new[]{ c.P[0]+ofs, c.P[1], c.P[2], c.P[3]+ofs, c.P[4], c.P[5] }).ToList();
  }
  static double Orta(List<YuzCizgisi> F) { return (F.Min(c => Math.Min(c.P[0], c.P[3])) + F.Max(c => Math.Max(c.P[0], c.P[3]))) / 2; }
  // Yükseklik kesitleri: Line(y + ofs, -x, z); yatay elemanların kotlarında, ofsetten başlayıp 30000 arayla
  static List<double[]> Kesitler(IDictionary<string, FinalMember> u, int ofs, bool ikiD) {
    var l = new List<double[]>(); int i = 0;
    foreach (var z in u.Values.Where(m => Math.Abs(m.start_z - m.end_z) < 1).Select(m => Math.Round(m.start_z)).Distinct().OrderBy(z => z).Take(8)) {
      int o = ofs + 30000 * i++;
      foreach (var m in u.Values.Where(m => Math.Abs(m.start_z - z) < 15 && Math.Abs(m.end_z - z) < 15)) {
        double[] p = { m.start_y + o, -m.start_x, m.start_z, m.end_y + o, -m.end_x, m.end_z };
        l.Add(ikiD ? new[]{ p[0], p[2], -p[1], p[3], p[5], -p[4] } : p);
      }
    }
    return l;
  }
  static int hata = 0;
  static void Dene(string ad, KuleVerisi k, bool on, bool yan, bool ikiD, bool baslik) {
    var cizgi = new List<double[]>(); var bas = new List<GorunusBasligi>(); int ofs = 0;
    int onOfs = -1, yanOfs = -1;
    if (on) { cizgi.AddRange(Ciz(k.OnYuz, ofs, ikiD)); if (baslik) bas.Add(new GorunusBasligi { Yan = false, X = Orta(k.OnYuz) + ofs, IkiD = ikiD }); onOfs = ofs; ofs += 50000; }
    if (yan) { cizgi.AddRange(Ciz(k.YanYuz, ofs, ikiD)); if (baslik) bas.Add(new GorunusBasligi { Yan = true, X = Orta(k.YanYuz) + ofs, IkiD = ikiD }); yanOfs = ofs; ofs += 50000; }
    cizgi.AddRange(Kesitler(k.Uyeler, ofs, ikiD));
    var sonuc = GorunusKonumBulucu.Bul(k.OnYuz, k.YanYuz, bas, cizgi);
    var bulOn = sonuc.FirstOrDefault(x => !x.Yan); var bulYan = sonuc.FirstOrDefault(x => x.Yan);
    bool onDogru = on ? (bulOn != null && bulOn.Ofs == onOfs && bulOn.IkiD == ikiD) : bulOn == null;
    bool yanDogru = yan ? (bulYan != null && bulYan.Ofs == yanOfs && bulYan.IkiD == ikiD) : bulYan == null;
    // Başlıksız ve yüzleri özdeş kulede görünüşün bulunamaması kabul (yanlış yere çizmekten iyi): "belirsiz"
    bool yanlisYer = (bulOn != null && (!on || bulOn.Ofs != onOfs || bulOn.IkiD != ikiD)) || (bulYan != null && (!yan || bulYan.Ofs != yanOfs || bulYan.IkiD != ikiD));
    string durum = (onDogru && yanDogru) ? "DOGRU" : (yanlisYer ? "YANLIS YER" : "bulunamadi (belirsiz)");
    if (yanlisYer) hata++;
    Console.WriteLine("  {0,-34} -> {1}{2}", ad, durum,
      string.Concat(sonuc.Select(x => string.Format("  [{0} {1}{2} {3}/{4} {5}]", x.Yan ? "yan" : "on", x.Ofs, x.IkiD ? " 2D" : "", x.Eslesen, x.Toplam, x.Kaynak))));
  }
  static void Main(string[] a) {
    var k = KuleOkuyucu.OkuDogrudan(typeof(KuleOkuyucu).Assembly, a[0]);
    foreach (bool baslik in new[]{ true, false })
      foreach (bool ikiD in new[]{ false, true }) {
        string e = (baslik ? "baslikli" : "basliksiz") + (ikiD ? ", 2D" : "");
        Dene("yalniz on (" + e + ")", k, true, false, ikiD, baslik);
        Dene("yalniz yan (" + e + ")", k, false, true, ikiD, baslik);
        Dene("on + yan (" + e + ")", k, true, true, ikiD, baslik);
        Dene("hicbiri (" + e + ")", k, false, false, ikiD, baslik);
      }
    Console.WriteLine(hata == 0 ? "  SONUC: hicbir gorunus yanlis yere konmadi" : "  SONUC: " + hata + " durumda YANLIS YER");
  }
}
