// Görünüş tamlık testi: her elemanın önden / yandan izdüşümü (PLS-TOWER görünüşü) çizimde
// (OutlineDrawing yüz çizgileri + eklentinin iç eleman izdüşümleri) var mı? Yan görünüşte uçtan görülen
// travers elemanları (gövde dışı) bilerek eklenmez; "yok" sayısı onlardır. Ayrıntı için ikinci argüman verin.
using System; using System.Linq; using System.Collections.Generic; using OtomatikKesit;
class P {
  static void Main(string[] a) {
    var k = KuleOkuyucu.OkuDogrudan(typeof(KuleOkuyucu).Assembly, a[0]);
    foreach (bool yan in new[]{false, true}) {
      var yuz = yan ? k.YanYuz : k.OnYuz;
      var ic = IcElemanPlani.Hesapla(k.Uyeler, yuz, yan);
      var cizilen = yuz.Concat(ic).Select(c => new[]{ c.P[0], c.P[2], c.P[3], c.P[5] }).ToList();
      int toplam = 0, var_ = 0; var eksik = new List<string>();
      foreach (var kv in k.Uyeler) {
        var m = kv.Value;
        double[] p = yan ? new[]{ m.start_y, -m.start_x, m.start_z, m.end_y, -m.end_x, m.end_z } : new[]{ m.start_x, m.start_y, m.start_z, m.end_x, m.end_y, m.end_z };
        double[] s = { p[0], p[2], p[3], p[5] };
        double L = Math.Sqrt((s[2]-s[0])*(s[2]-s[0])+(s[3]-s[1])*(s[3]-s[1]));
        if (L < 20) continue;          // bakış yönünde: nokta
        toplam++;
        if (IcElemanPlani.Kapsaniyor(s, cizilen)) var_++;
        else eksik.Add(string.Format("{0} ({1}) z {2:0}-{3:0} derinlik {4:0}/{5:0}", kv.Key, m.group_label, p[2], p[5], p[1], p[4]));
      }
      Console.WriteLine("  {0}: izdusum {1}, cizimde var {2}, yok {3}", yan ? "yan" : "on", toplam, var_, eksik.Count);
      foreach (var e in eksik.Take(a.Length > 1 ? 1000 : 0)) Console.WriteLine("     yok: " + e);
    }
  }
}
