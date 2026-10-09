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
    var beklenen = new Dictionary<string,int>(); char L = 'A'; int ri = 0;
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
      List<string> keys; KesitSecici.Sec(uyeler, tanim, out keys); beklenen[ad] = keys.Count;
    }
    t.GetMethod("SectionDetection", H).Invoke(o, new object[]{grid, a[0], 15});
    var section = (IDictionary)t.GetProperty("section").GetValue(o, null);
    var say = new Dictionary<string,int>();
    foreach (DictionaryEntry de in section) { string s = (string)de.Value.GetType().GetProperty("s_name").GetValue(de.Value, null); int n; say.TryGetValue(s, out n); say[s] = n + 1; }
    int eksik = 0;
    foreach (var kv in beklenen) { int n; say.TryGetValue(kv.Key, out n); if (n != kv.Value) eksik++;
      Console.WriteLine("  {0}: beklenen {1,3}, cizilecek {2,3}{3}", kv.Key, kv.Value, n, n != kv.Value ? "   <-- EKSIK " + (kv.Value - n) : ""); }
    Console.WriteLine(eksik == 0 ? "  SONUC: butun kesitler tam" : "  SONUC: " + eksik + " kesit eksik");
    return eksik;
  }
}
