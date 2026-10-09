// OutlineDrawing'in GERCEK SectionDetection kodunu (sahte AutoCAD geometri kutuphaneleriyle) calistirir;
// her kesitte cizilecek eleman sayisini OutlineDrawing secim kuralinin bekledigiyle karsilastirir.
//   bash test/kesit_testi.sh /yol/OutlineDrawing.dll kule.tow
using System; using System.Collections; using System.Collections.Generic; using System.Linq; using System.Reflection; using OtomatikKesit;
public static class SD { const BindingFlags H = BindingFlags.Instance|BindingFlags.Public|BindingFlags.NonPublic;
  public static int Main(string[] a) {
    var asm = typeof(KuleOkuyucu).Assembly; var t = asm.GetType("OutlineDrawing.dataprocessing");
    var o = Activator.CreateInstance(t, true);
    t.GetMethod("datapreprocessing", H).Invoke(o, new object[]{a[0]});
    var uyeler = KuleOkuyucu.Kopyala(o);
    AutoSectionReport r; var secs = AutoSectionDetector.DetectHidden(uyeler, KuleOkuyucu.Gorunenler(o), new AutoSectionOptions(), out r);
    // grid -> Dictionary<string, SectionRowData>
    Type rowT = asm.GetType("OutlineDrawing.SectionRowData"), ptT = asm.GetType("OutlineDrawing.SectionPoint3D");
    var dictT = typeof(Dictionary<,>).MakeGenericType(typeof(string), rowT);
    var grid = (IDictionary)Activator.CreateInstance(dictT);
    var beklenen = new Dictionary<string,int>(); var tanimlar = new Dictionary<string, KesitTanimi>(); char L = 'A'; int ri = 0;
    foreach (var c in secs) {
      string ad = (L++).ToString();
      object row = Activator.CreateInstance(rowT);
      rowT.GetProperty("SectionName").SetValue(row, ad, null);
      rowT.GetProperty("RowIndex").SetValue(row, ri++, null);
      var pts = (IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(ptT));
      var tanim = new KesitTanimi();
      if (c.Kind == AutoSectionKind.Height) { rowT.GetProperty("Height").SetValue(row, (double?)c.Height.Value, null); tanim.Yukseklik = c.Height; }
      else foreach (var p in c.Points) { pts.Add(Activator.CreateInstance(ptT, p.X, p.Y, p.Z)); tanim.Noktalar.Add(new[]{p.X,p.Y,p.Z}); }
      rowT.GetProperty("Points").SetValue(row, pts, null);
      grid[ad] = row;
      List<string> keys; KesitSecici.Sec(uyeler, tanim, out keys); beklenen[ad] = keys.Count; tanimlar[ad] = tanim;
    }
    t.GetMethod("SectionDetection", H).Invoke(o, new object[]{grid, a[0], 15});
    var section = (IDictionary)t.GetProperty("section").GetValue(o, null);
    // Çizim: yükseklik kesiti Line(s.y+ofs, -s.x, s.z); nokta kesiti Line(s.x+ofs, s.y, s.z)  -> (u, v), derinlik w
    var cizim = new Dictionary<string, List<double[]>>(); var cizimAd = new Dictionary<string, List<string>>(); var derin = new Dictionary<string, double[]>();
    foreach (DictionaryEntry de in section) {
      object sv = de.Value; Type st = sv.GetType();
      Func<string,double> D = ad => Convert.ToDouble(st.GetProperty(ad).GetValue(sv, null));
      string s = (string)st.GetProperty("s_name").GetValue(sv, null);
      bool yuk = Convert.ToInt32(st.GetProperty("Type").GetValue(sv, null)) == 0;
      double[] seg = yuk ? new[]{ D("start_y"), D("start_z"), D("end_y"), D("end_z") } : new[]{ D("start_x"), D("start_z"), D("end_x"), D("end_z") };
      double w1 = yuk ? -D("start_x") : D("start_y"), w2 = yuk ? -D("end_x") : D("end_y");
      if (!cizim.ContainsKey(s)) { cizim[s] = new List<double[]>(); cizimAd[s] = new List<string>(); derin[s] = new[]{ double.MaxValue, double.MinValue }; }
      cizim[s].Add(seg); cizimAd[s].Add((string)st.GetProperty("name").GetValue(sv, null));
      derin[s][0] = Math.Min(derin[s][0], Math.Min(w1, w2)); derin[s][1] = Math.Max(derin[s][1], Math.Max(w1, w2));
    }
    var red = KuleOkuyucu.Redundantlar(o, a[0]);
    int stilHata = 0, eslesmeyen = 0;
    foreach (var kv in cizim) {
      var tanim = tanimlar[kv.Key]; List<string> anahtar; var kaynak = KesitSecici.Sec(uyeler, tanim, out anahtar);
      var es = Eslestirici.Eslestir(kaynak, kv.Value, 2.0);
      for (int i = 0; i < es.Length; i++) {
        if (es[i] < 0) { eslesmeyen++; continue; }
        bool gercek = red.Contains(uyeler[cizimAd[kv.Key][i]].group_label), bulunan = red.Contains(uyeler[anahtar[es[i]]].group_label);
        if (gercek != bulunan) stilHata++;
      }
      Console.WriteLine("  {0}: derinlik {1:0}..{2:0} mm, eslesen {3}/{4}", kv.Key, derin[kv.Key][0], derin[kv.Key][1], es.Count(x => x >= 0), es.Length);
    }
    Console.WriteLine("  Cizim eslestirme: eslesmeyen {0}, yanlis stil {1}", eslesmeyen, stilHata);
    var say = new Dictionary<string,int>();
    foreach (DictionaryEntry de in section) { string s = (string)de.Value.GetType().GetProperty("s_name").GetValue(de.Value, null); int n; say.TryGetValue(s, out n); say[s] = n + 1; }
    int eksik = 0;
    foreach (var kv in beklenen) { int n; say.TryGetValue(kv.Key, out n); if (n != kv.Value) eksik++;
      Console.WriteLine("  {0}: beklenen {1,3}, cizilecek {2,3}{3}", kv.Key, kv.Value, n, n != kv.Value ? "   <-- EKSIK " + (kv.Value - n) : ""); }
    Console.WriteLine(eksik == 0 ? "  SONUC: butun kesitler tam" : "  SONUC: " + eksik + " kesit eksik");
    return eksik + stilHata + eslesmeyen;
  }
}
