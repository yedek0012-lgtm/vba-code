// Ayna parçası testi (sentetik): bir kotta 4 kopuk köşe paneli, biri öbürlerinin aynası DEĞİL (köşegen ters yönde).
// Otomatik kesit yalnız "tek ayna parça"yı kesit alırsa ters köşegen hiçbir yerde çizilmez (1.21'e kadar: AynaGrup
// yalnız orta noktaları karşılaştırıyordu; iki köşegenin orta noktası aynı). Gerçek SectionDetection ile denetlenir.
//   bash test/ayna_parca_testi.sh eklenti/bin/tek/OutlineDrawing.dll
using System; using System.Collections; using System.Collections.Generic; using System.Linq; using System.Reflection; using System.Globalization; using OtomatikKesit;
public static class AynaParcaTesti {
  const BindingFlags H = BindingFlags.Instance|BindingFlags.Static|BindingFlags.Public|BindingFlags.NonPublic;
  static Dictionary<string, FinalMember> M = new Dictionary<string, FinalMember>();
  static void Ekle(string ad, double x1, double y1, double z1, double x2, double y2, double z2, string g) {
    M[ad] = new FinalMember { start_x=x1, start_y=y1, start_z=z1, end_x=x2, end_y=y2, end_z=z2, group_label=g, size="L50x5", material="S235", section_label="BODY" };
  }
  public static int Main(string[] a) {
    double z = 10000;
    // 4 kopuk köşe paneli (bacak köşesi plan bağlantısı gibi): her biri 4 kenar + 1 köşegen.
    int[] sx = { 1, -1, -1, 1 }, sy = { -1, -1, 1, 1 };
    for (int c = 0; c < 4; c++) {
      double x1 = sx[c]*1000, x2 = sx[c]*2000, y1 = sy[c]*1000, y2 = sy[c]*2000;
      Ekle("K"+c+"a", x1,y1,z, x2,y1,z, "PK"); Ekle("K"+c+"b", x2,y1,z, x2,y2,z, "PK");
      Ekle("K"+c+"c", x2,y2,z, x1,y2,z, "PK"); Ekle("K"+c+"d", x1,y2,z, x1,y1,z, "PK");
      if (c == 2) Ekle("K2_TERS", x1,y2,z, x2,y1,z, "PDT");   // -X/+Y köşesinde köşegen öbür yönde (ayna değil)
      else        Ekle("K"+c+"_DIAG", x1,y1,z, x2,y2,z, "PD");
    }
    AutoSectionReport r;
    var secs = AutoSectionDetector.DetectHidden(M, new List<string>{"yok"}, new AutoSectionOptions(), out r);
    Console.WriteLine("Otomatik kesit: {0}; görünmeyen {1}, kapsandı denilen {2}, kapsanmayan rapor {3}", secs.Count, r.HiddenMembers, r.CoveredHidden, r.UncoveredKeys.Count);
    foreach (var c in secs) Console.WriteLine("  " + c.Kind + " " + c.Description + " | noktalar: " + string.Join("  ", c.Points.Select(AutoSectionDetector.FormatPoint)));
    // Gerçek OutlineDrawing SectionDetection
    var asm = typeof(KuleOkuyucu).Assembly; var t = asm.GetType("OutlineDrawing.dataprocessing");
    var o = Activator.CreateInstance(t, true);
    var fmT = asm.GetType("OutlineDrawing.FinalMember");
    var fm = (IDictionary)Activator.CreateInstance(typeof(Dictionary<,>).MakeGenericType(typeof(string), fmT));
    foreach (var kv in M) { var v = Activator.CreateInstance(fmT); foreach (var p in new[]{"start_x","start_y","start_z","end_x","end_y","end_z"}) fmT.GetProperty(p).SetValue(v, typeof(FinalMember).GetProperty(p).GetValue(kv.Value, null), null);
      fmT.GetProperty("group_label").SetValue(v, kv.Value.group_label, null); fmT.GetProperty("size").SetValue(v, "L50x5", null); fmT.GetProperty("material").SetValue(v, "S235", null); fmT.GetProperty("section_label").SetValue(v, "BODY", null); fm[kv.Key] = v; }
    t.GetProperty("finalMember").SetValue(o, fm, null);
    var secT = t.GetProperty("section").PropertyType; t.GetProperty("section").SetValue(o, Activator.CreateInstance(secT), null);
    Type rowT = asm.GetType("OutlineDrawing.SectionRowData"), ptT = asm.GetType("OutlineDrawing.SectionPoint3D");
    var grid = (IDictionary)Activator.CreateInstance(typeof(Dictionary<,>).MakeGenericType(typeof(string), rowT));
    int ri = 0;
    foreach (var c in secs) { string ad = ((char)('A'+ri)).ToString(); object row = Activator.CreateInstance(rowT);
      rowT.GetProperty("SectionName").SetValue(row, ad, null); rowT.GetProperty("RowIndex").SetValue(row, ri++, null);
      var pts = (IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(ptT));
      if (c.Kind == AutoSectionKind.Height) rowT.GetProperty("Height").SetValue(row, (double?)c.Height.Value, null);
      else foreach (var p in c.Points) pts.Add(Activator.CreateInstance(ptT, p.X, p.Y, p.Z));
      rowT.GetProperty("Points").SetValue(row, pts, null); grid[ad] = row; }
    string yol = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ayna_parca_kule.tow");
    t.GetMethod("SectionDetection", H).Invoke(o, new object[]{grid, yol, 15});
    var cizilen = new HashSet<string>();
    foreach (DictionaryEntry de in (IDictionary)t.GetProperty("section").GetValue(o, null)) { string k = (string)de.Key; int i = k.IndexOf(KesitAnahtari.Ayrac); cizilen.Add(i >= 0 ? k.Substring(i+1) : k); }
    Console.WriteLine("Gerçek SectionDetection'ın çizdiği elemanlar: " + string.Join(",", cizilen.OrderBy(x => x)));
    // Ayna ailesi: elemanın kendisi veya X/Y/XY aynası çizildi mi?
    Func<double,string> R = v => (Math.Round(v)+0.0).ToString(CultureInfo.InvariantCulture);
    Func<FinalMember,double,double,string> G = (m,fx,fy) => { string p = R(fx*m.start_x)+","+R(fy*m.start_y)+","+R(m.start_z), q = R(fx*m.end_x)+","+R(fy*m.end_y)+","+R(m.end_z); return string.CompareOrdinal(p,q)<0? p+"|"+q : q+"|"+p; };
    var cizilenGeo = new HashSet<string>(cizilen.Select(k => G(M[k],1,1)));
    int yok = 0;
    foreach (var kv in M) {
      bool aile = new[]{ new[]{1.0,1.0}, new[]{-1.0,1.0}, new[]{1.0,-1.0}, new[]{-1.0,-1.0} }.Any(s => cizilenGeo.Contains(G(kv.Value, s[0], s[1])));
      if (!aile) { yok++; Console.WriteLine("  HİÇBİR KESİTTE YOK (aynası da yok): " + kv.Key + " " + kv.Value.group_label); }
    }
    Console.WriteLine(yok == 0 ? "SONUC: ayna parca testi TAMAM" : "SONUC: ayna parca testi HATA");
    return yok;
  }
}
