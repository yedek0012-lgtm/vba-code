// =====================================================================
//  AutoSectionDetector.cs  -  OutlineDrawing için OTOMATİK KESİT bulucu
// ---------------------------------------------------------------------
//  OutlineDrawing projesine (namespace OutlineDrawing) olduğu gibi eklenir.
//  AutoCAD API'sine bağımlı DEĞİLDİR; yalnızca .tow okunduktan sonra oluşan
//  dataprocessing.finalMember sözlüğü (Dictionary<string, FinalMember>)
//  üzerinde çalışır.
//
//  Bulduğu kesitler mevcut grid / SectionRowData formatıyla birebir aynıdır:
//    - Yatay kesit  -> ColumnHeight  (Z, mm)
//    - Eğik düzlem  -> ColumnFirstPoint..ColumnFourthPoint ("x;y;z")
//  Yani mevcut CollectSectionsFromGrid -> SectionDetection -> DrawSections
//  akışı hiç değişmeden kullanılır; bu sınıf yalnızca "nereden kesit
//  alınmalı" sorusunu cevaplar.
//
//  Koordinatlar finalMember ile aynıdır (mm, RotateAllMinus90AboutZ sonrası),
//  yani 3D modelde elle tıklanan noktalarla aynı sistemdedir.
// =====================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OutlineDrawing
{
    public enum AutoSectionKind { Height = 0, Plane = 1 }

    public class AutoSectionOptions
    {
        // ---- Yatay (yükseklik) kesitleri --------------------------------
        /// <summary>mm. runtowerfile içindeki SectionDetection(tol = 15) ile aynı olmalı.</summary>
        public double HeightTolerance = 15.0;
        /// <summary>mm. Eleman bu mesafe içinde dış konturun (convex hull) kenarındaysa "yüz elemanı" sayılır.</summary>
        public double PerimeterTolerance = 10.0;
        /// <summary>Seviyede en az bu kadar iç (plan) eleman varsa kesit önerilir. 0 = yatay elemanı olan her seviye.</summary>
        public int MinPlanMembers = 1;
        /// <summary>Seviyede en az bu kadar yatay eleman olmalı.</summary>
        public int MinLevelMembers = 3;

        // ---- Eğik düzlem (nokta) kesitleri ------------------------------
        public bool DetectPlanes = true;
        /// <summary>mm. Eleman iki ucunun da düzleme uzaklığı bundan küçükse düzlemdedir.</summary>
        public double PlaneTolerance = 20.0;
        /// <summary>Derece. Bundan az eğimli düzlem yatay sayılır (yükseklik kesiti zaten yakalar).</summary>
        public double MinPlaneTiltDeg = 3.0;
        /// <summary>Derece. Bundan dik düzlemler yüz düzlemidir (ön/yan görünüşte zaten var).</summary>
        public double MaxPlaneTiltDeg = 40.0;
        /// <summary>Düzlemdeki bağlı eleman grubunun en az eleman sayısı.</summary>
        public int MinPlaneMembers = 4;
        /// <summary>mm. Kesit dörtgeni bu kadar dışa büyütülür (PointInPolygon kenarda kaçırmasın).</summary>
        public double PolygonMargin = 50.0;

        // ---- Tekrarlar ---------------------------------------------------
        /// <summary>±X / ±Y ayna simetriği olan düzlem kesitlerinden yalnız birini tut (+X / -Y tarafı).</summary>
        public bool MergeSymmetric = true;
        /// <summary>Aynı grup listesine sahip yatay seviyelerden yalnız ilkini tut ("tipik kesit").</summary>
        public bool MergeIdenticalLevels = false;
        /// <summary>mm. Simetri eşleştirmesinde ağırlık merkezi toleransı.</summary>
        public double SymmetryTolerance = 100.0;

        /// <summary>true: A kesiti en üstte, aşağı doğru B, C, ...</summary>
        public bool TopDown = true;
    }

    public class AutoSectionCandidate
    {
        public AutoSectionKind Kind;
        /// <summary>Kind == Height ise kesit kotu (mm).</summary>
        public double? Height;
        /// <summary>Kind == Plane ise 4 köşe noktası (grid'deki First..Fourth Point).</summary>
        public List<SectionPoint3D> Points = new List<SectionPoint3D>();
        /// <summary>Kesite giren finalMember anahtarları.</summary>
        public List<string> MemberKeys = new List<string>();
        /// <summary>Yatay kesitte: dış kontur üzerinde olmayan (plan) eleman sayısı.</summary>
        public int PlanMemberCount;
        /// <summary>Düzlem kesitte: düzlemin yataydan eğimi (derece).</summary>
        public double TiltDeg;
        /// <summary>Sıralama kotu (mm).</summary>
        public double SortZ;
        public string Description = "";
    }

    public static class AutoSectionDetector
    {
        // =================================================================
        //  Giriş noktası
        // =================================================================
        public static List<AutoSectionCandidate> Detect(IDictionary<string, FinalMember> members, AutoSectionOptions opt = null)
        {
            if (opt == null) opt = new AutoSectionOptions();
            var result = new List<AutoSectionCandidate>();
            if (members == null || members.Count == 0) return result;

            var bars = new List<Bar>();
            foreach (var kv in members)
            {
                if (kv.Value == null) continue;
                var b = new Bar(kv.Key, kv.Value);
                if (b.Length < 1.0) continue;
                bars.Add(b);
            }

            result.AddRange(DetectHeightSections(bars, opt));
            if (opt.DetectPlanes) result.AddRange(DetectPlaneSections(bars, opt));

            if (opt.TopDown)
                result = result.OrderByDescending(c => c.SortZ).ThenBy(c => (int)c.Kind).ToList();
            else
                result = result.OrderBy(c => c.SortZ).ThenBy(c => (int)c.Kind).ToList();
            return result;
        }

        /// <summary>Grid hücresine yazılacak nokta metni (btnSelectMultiplePoint ile aynı biçim).</summary>
        public static string FormatPoint(SectionPoint3D p)
        {
            return string.Format("{0};{1};{2}",
                p.X.ToString("0.###", CultureInfo.CurrentCulture),
                p.Y.ToString("0.###", CultureInfo.CurrentCulture),
                p.Z.ToString("0.###", CultureInfo.CurrentCulture));
        }

        // =================================================================
        //  1) YATAY KESİTLER
        //     Yatay elemanları kota göre grupla; her seviyede dış konturda
        //     (convex hull kenarında) OLMAYAN eleman varsa = plan çaprazı /
        //     diyafram / travers alt düzlemi vardır -> kesit gerekir.
        //     Sadece dış kontur elemanları olan seviye ön/yan görünüşte
        //     zaten görünür -> kesit gerekmez.
        // =================================================================
        private static List<AutoSectionCandidate> DetectHeightSections(List<Bar> bars, AutoSectionOptions opt)
        {
            var list = new List<AutoSectionCandidate>();
            double tol = opt.HeightTolerance;

            var horiz = bars.Where(b => Math.Abs(b.A.Z - b.B.Z) <= tol)
                            .OrderBy(b => b.Mid.Z).ToList();
            if (horiz.Count == 0) return list;

            // Kot kümeleri: kümenin ilk elemanından en fazla tol uzaktaki elemanlar
            var levels = new List<double>();
            int i = 0;
            while (i < horiz.Count)
            {
                double z0 = horiz[i].Mid.Z, sum = 0; int n = 0;
                while (i < horiz.Count && horiz[i].Mid.Z - z0 <= tol) { sum += horiz[i].Mid.Z; n++; i++; }
                levels.Add(sum / n);
            }

            var seenSignatures = new Dictionary<string, double>();
            foreach (double h in levels)
            {
                // SectionDetection ile birebir aynı seçim kuralı
                var lvl = bars.Where(b => Math.Abs(b.A.Z - h) <= tol && Math.Abs(b.B.Z - h) <= tol).ToList();
                if (lvl.Count < opt.MinLevelMembers) continue;

                // Kontur her bağlı grup için ayrı: travers üzerindeki tek bir yatay eleman
                // tesadüfen gövde kotuna denk gelirse gövde kuşakları "iç eleman" sayılmasın.
                int planCount = 0;
                foreach (var grp in ConnectedGroups(lvl, true))
                {
                    if (grp.Count < 3) continue;
                    var pts = new List<V2>();
                    foreach (var b in grp) { pts.Add(new V2(b.A.X, b.A.Y)); pts.Add(new V2(b.B.X, b.B.Y)); }
                    var hull = ConvexHull(pts);
                    if (hull.Count < 3) continue;
                    foreach (var b in grp)
                        if (!OnHullEdge(new V2(b.A.X, b.A.Y), new V2(b.B.X, b.B.Y), hull, opt.PerimeterTolerance))
                            planCount++;
                }
                if (planCount < opt.MinPlanMembers) continue;

                string sig = GroupSignature(lvl);
                string note = "";
                double prevZ;
                if (seenSignatures.TryGetValue(sig, out prevZ))
                {
                    if (opt.MergeIdenticalLevels) continue;
                    note = string.Format(CultureInfo.InvariantCulture, " (Z={0:0} ile aynı gruplar)", prevZ);
                }
                else seenSignatures[sig] = h;

                double hr = Math.Round(h, 3);
                list.Add(new AutoSectionCandidate
                {
                    Kind = AutoSectionKind.Height,
                    Height = hr,
                    SortZ = hr,
                    PlanMemberCount = planCount,
                    MemberKeys = lvl.Select(b => b.Key).ToList(),
                    Description = string.Format(CultureInfo.InvariantCulture,
                        "Yatay kesit Z={0:0} mm: {1} yatay eleman, {2} plan elemanı{3}", hr, lvl.Count, planCount, note)
                });
            }
            return list;
        }

        // =================================================================
        //  2) EĞİK DÜZLEM KESİTLERİ (travers üst başlık düzlemi, eğik
        //     diyaframlar vb.)
        //     Ortak düğümü olan "yatıkça" eleman çiftlerinden aday düzlemler
        //     üretilir; yataydan MinTilt..MaxTilt eğimli olanlar tutulur.
        //     Düzlemdeki elemanlar bağlı gruplara ayrılır; içinde en az bir
        //     üçgen (gerçek kafes) olan gruplar kesit olur.
        // =================================================================
        private static List<AutoSectionCandidate> DetectPlaneSections(List<Bar> bars, AutoSectionOptions opt)
        {
            var list = new List<AutoSectionCandidate>();
            double sinMax = Math.Sin(opt.MaxPlaneTiltDeg * Math.PI / 180.0);
            double sinParallel = Math.Sin(10.0 * Math.PI / 180.0);

            // Eğimi MaxTilt'ten büyük bir eleman MaxTilt'ten yatık bir düzlemde olamaz
            var flat = bars.Where(b => Math.Abs(b.Dir.Z) <= sinMax + 1e-9).ToList();
            if (flat.Count < opt.MinPlaneMembers) return list;

            var nodeBars = new Dictionary<string, List<Bar>>();
            var nodePos = new Dictionary<string, V3>();
            foreach (var b in flat)
            {
                AddNode(nodeBars, nodePos, b.KeyA, b.A, b);
                AddNode(nodeBars, nodePos, b.KeyB, b.B, b);
            }

            // ---- aday düzlemler
            var planeKeys = new HashSet<string>();
            var planes = new List<PlaneCand>();
            foreach (var kv in nodeBars)
            {
                var lst = kv.Value;
                V3 p = nodePos[kv.Key];
                for (int a = 0; a < lst.Count; a++)
                    for (int c = a + 1; c < lst.Count; c++)
                    {
                        V3 cr = V3.Cross(lst[a].Dir, lst[c].Dir);
                        double len = cr.Length;
                        if (len < sinParallel) continue;           // paralel
                        V3 n = cr / len;
                        if (n.Z < 0) n = -n;
                        double tilt = Math.Acos(Math.Min(1.0, n.Z)) * 180.0 / Math.PI;
                        if (tilt < opt.MinPlaneTiltDeg || tilt > opt.MaxPlaneTiltDeg) continue;
                        double d = V3.Dot(n, p);
                        string key = Math.Round(n.X * 100).ToString(CultureInfo.InvariantCulture) + "|" +
                                     Math.Round(n.Y * 100).ToString(CultureInfo.InvariantCulture) + "|" +
                                     Math.Round(d / opt.PlaneTolerance).ToString(CultureInfo.InvariantCulture);
                        if (!planeKeys.Add(key)) continue;
                        planes.Add(new PlaneCand { N = n, D = d, Tilt = tilt });
                    }
            }

            // ---- düzlemdeki bağlı eleman grupları
            var comps = new List<PlaneComp>();
            var compKeys = new HashSet<string>();
            foreach (var pl in planes)
            {
                var inPlane = flat.Where(b => Math.Abs(V3.Dot(pl.N, b.A) - pl.D) <= opt.PlaneTolerance &&
                                              Math.Abs(V3.Dot(pl.N, b.B) - pl.D) <= opt.PlaneTolerance).ToList();
                if (inPlane.Count < opt.MinPlaneMembers) continue;

                foreach (var comp in ConnectedGroups(inPlane))
                {
                    if (comp.Count < opt.MinPlaneMembers) continue;
                    if (!SpansTwoDirections(comp, sinParallel)) continue;
                    if (!HasTriangle(comp)) continue;              // kafes değilse (yüzleri kesen tesadüfi düzlem) at
                    string ck = string.Join("|", comp.Select(b => b.Key).OrderBy(s => s, StringComparer.Ordinal));
                    if (!compKeys.Add(ck)) continue;
                    comps.Add(new PlaneComp { Plane = pl, Bars = comp });
                }
            }

            // ---- alt küme / neredeyse aynı grupları ele (büyükten küçüğe)
            comps = comps.OrderByDescending(c => c.Bars.Count).ToList();
            var kept = new List<PlaneComp>();
            foreach (var c in comps)
            {
                var keys = new HashSet<string>(c.Bars.Select(b => b.Key));
                bool covered = kept.Any(k => keys.Count(x => k.KeySet.Contains(x)) >= 0.8 * keys.Count);
                if (covered) continue;
                c.KeySet = keys;
                c.Centroid = Centroid(c.Bars);
                kept.Add(c);
            }

            // ---- ayna simetriği olanları birleştir (+X ve -Y tarafı tercih)
            double st = opt.SymmetryTolerance;
            var ordered = kept.OrderBy(c => c.Centroid.X < -st ? 1 : 0)
                              .ThenBy(c => c.Centroid.Y > st ? 1 : 0)
                              .ThenByDescending(c => c.Bars.Count).ToList();
            var final = new List<PlaneComp>();
            foreach (var c in ordered)
            {
                if (opt.MergeSymmetric && final.Any(k => IsMirror(k, c, st))) continue;
                final.Add(c);
            }

            foreach (var c in final)
            {
                var corners = BoundingQuad(c, opt.PolygonMargin);
                if (corners == null) continue;
                list.Add(new AutoSectionCandidate
                {
                    Kind = AutoSectionKind.Plane,
                    Points = corners,
                    MemberKeys = c.Bars.Select(b => b.Key).ToList(),
                    TiltDeg = Math.Round(c.Plane.Tilt, 1),
                    SortZ = c.Centroid.Z,
                    Description = string.Format(CultureInfo.InvariantCulture,
                        "Eğik düzlem kesiti: eğim {0:0.0}°, {1} eleman, merkez ({2:0}; {3:0}; {4:0})",
                        c.Plane.Tilt, c.Bars.Count, c.Centroid.X, c.Centroid.Y, c.Centroid.Z)
                });
            }
            return list;
        }

        // =================================================================
        //  Yardımcılar
        // =================================================================
        private static void AddNode(Dictionary<string, List<Bar>> nodeBars, Dictionary<string, V3> nodePos, string key, V3 p, Bar b)
        {
            List<Bar> l;
            if (!nodeBars.TryGetValue(key, out l)) { l = new List<Bar>(); nodeBars[key] = l; nodePos[key] = p; }
            l.Add(b);
        }

        private static string NodeKey(V3 p)
        {
            return Math.Round(p.X).ToString(CultureInfo.InvariantCulture) + "," +
                   Math.Round(p.Y).ToString(CultureInfo.InvariantCulture) + "," +
                   Math.Round(p.Z).ToString(CultureInfo.InvariantCulture);
        }

        private static string NodeKeyXY(V3 p)
        {
            return Math.Round(p.X).ToString(CultureInfo.InvariantCulture) + "," +
                   Math.Round(p.Y).ToString(CultureInfo.InvariantCulture);
        }

        private static string GroupSignature(IEnumerable<Bar> bars)
        {
            return string.Join(",", bars.Select(b => b.Group ?? "").OrderBy(s => s, StringComparer.Ordinal));
        }

        private static V3 Centroid(List<Bar> bars)
        {
            V3 s = new V3(0, 0, 0);
            foreach (var b in bars) s = s + b.Mid;
            return s / bars.Count;
        }

        /// <summary>b, a'nın X=0 / Y=0 / her ikisine göre aynası mı? (eleman orta noktaları birebir eşleşmeli;
        /// grup adları farklı olsa da geometri aynıysa aynı kesittir)</summary>
        private static bool IsMirror(PlaneComp a, PlaneComp b, double tol)
        {
            if (a.Bars.Count != b.Bars.Count) return false;
            int[,] signs = { { -1, 1 }, { 1, -1 }, { -1, -1 } };
            for (int s = 0; s < 3; s++)
            {
                double fx = signs[s, 0], fy = signs[s, 1];
                V3 mc = new V3(fx * b.Centroid.X, fy * b.Centroid.Y, b.Centroid.Z);
                if ((mc - a.Centroid).Length > tol) continue;
                bool all = b.Bars.All(bb =>
                {
                    V3 m = new V3(fx * bb.Mid.X, fy * bb.Mid.Y, bb.Mid.Z);
                    return a.Bars.Any(ab => (ab.Mid - m).Length <= tol);
                });
                if (all) return true;
            }
            return false;
        }

        /// <summary>Ortak düğümü olan elemanları gruplar. planXY = true: düğüm yalnız (x,y) ile eşlenir
        /// (aynı seviyede, tolerans içindeki küçük Z farkları bağlantıyı koparmasın).</summary>
        private static List<List<Bar>> ConnectedGroups(List<Bar> bars, bool planXY = false)
        {
            Func<Bar, string> ka = b => planXY ? b.KeyAxy : b.KeyA;
            Func<Bar, string> kb = b => planXY ? b.KeyBxy : b.KeyB;
            var parent = new Dictionary<string, string>();
            Func<string, string> find = null;
            find = k =>
            {
                string r = k;
                while (parent[r] != r) r = parent[r];
                while (parent[k] != r) { string nx = parent[k]; parent[k] = r; k = nx; }
                return r;
            };
            foreach (var b in bars)
            {
                string a = ka(b), c = kb(b);
                if (!parent.ContainsKey(a)) parent[a] = a;
                if (!parent.ContainsKey(c)) parent[c] = c;
                string ra = find(a), rb = find(c);
                if (ra != rb) parent[ra] = rb;
            }
            return bars.GroupBy(b => find(ka(b))).Select(g => g.ToList()).ToList();
        }

        private static bool SpansTwoDirections(List<Bar> bars, double sinParallel)
        {
            V3 d0 = bars[0].Dir;
            return bars.Any(b => V3.Cross(d0, b.Dir).Length >= sinParallel);
        }

        private static bool HasTriangle(List<Bar> bars)
        {
            var adj = new Dictionary<string, HashSet<string>>();
            foreach (var b in bars)
            {
                if (b.KeyA == b.KeyB) continue;
                HashSet<string> s;
                if (!adj.TryGetValue(b.KeyA, out s)) { s = new HashSet<string>(); adj[b.KeyA] = s; }
                s.Add(b.KeyB);
                if (!adj.TryGetValue(b.KeyB, out s)) { s = new HashSet<string>(); adj[b.KeyB] = s; }
                s.Add(b.KeyA);
            }
            foreach (var b in bars)
            {
                HashSet<string> sa, sb;
                if (!adj.TryGetValue(b.KeyA, out sa) || !adj.TryGetValue(b.KeyB, out sb)) continue;
                foreach (var w in sa)
                    if (w != b.KeyB && sb.Contains(w)) return true;
            }
            return false;
        }

        /// <summary>
        /// Düzlemdeki grubu saran minimum alanlı dikdörtgen (dışa margin kadar büyütülmüş).
        /// Köşeler tam düzlem üzerindedir; BuildPlaneFromPoints (P0, P1, Pson) aynı düzlemi kurar.
        /// </summary>
        private static List<SectionPoint3D> BoundingQuad(PlaneComp c, double margin)
        {
            V3 n = c.Plane.N;
            V3 u = V3.Cross(new V3(0, 0, 1), n);              // düzlemdeki yatay doğrultu
            if (u.Length < 1e-9) return null;
            u = u / u.Length;
            V3 v = V3.Cross(n, u); v = v / v.Length;          // düzlemde eğim yönü
            V3 o = c.Centroid - n * (V3.Dot(n, c.Centroid) - c.Plane.D);

            var p2 = new List<V2>();
            foreach (var b in c.Bars)
            {
                p2.Add(new V2(V3.Dot(b.A - o, u), V3.Dot(b.A - o, v)));
                p2.Add(new V2(V3.Dot(b.B - o, u), V3.Dot(b.B - o, v)));
            }
            var hull = ConvexHull(p2);
            if (hull.Count < 3) return null;

            double bestArea = double.MaxValue, bx = 0, by = 0, minA = 0, maxA = 0, minB = 0, maxB = 0;
            for (int i = 0; i < hull.Count; i++)
            {
                V2 e = hull[(i + 1) % hull.Count] - hull[i];
                double el = e.Length;
                if (el < 1e-9) continue;
                double ex = e.X / el, ey = e.Y / el;
                double a0 = double.MaxValue, a1 = double.MinValue, b0 = double.MaxValue, b1 = double.MinValue;
                foreach (var q in hull)
                {
                    double a = q.X * ex + q.Y * ey, bb = -q.X * ey + q.Y * ex;
                    a0 = Math.Min(a0, a); a1 = Math.Max(a1, a); b0 = Math.Min(b0, bb); b1 = Math.Max(b1, bb);
                }
                double area = (a1 - a0) * (b1 - b0);
                if (area < bestArea) { bestArea = area; bx = ex; by = ey; minA = a0; maxA = a1; minB = b0; maxB = b1; }
            }
            minA -= margin; maxA += margin; minB -= margin; maxB += margin;

            var res = new List<SectionPoint3D>();
            double[,] ab = { { minA, minB }, { maxA, minB }, { maxA, maxB }, { minA, maxB } };
            for (int k = 0; k < 4; k++)
            {
                double a = ab[k, 0], bb = ab[k, 1];
                double x2 = a * bx - bb * by, y2 = a * by + bb * bx;     // dikdörtgen eksenlerinden (u,v)'ye
                V3 p = o + u * x2 + v * y2;
                res.Add(new SectionPoint3D(Math.Round(p.X, 3), Math.Round(p.Y, 3), Math.Round(p.Z, 3)));
            }
            return res;
        }

        private static List<V2> ConvexHull(List<V2> pts)
        {
            var p = pts.OrderBy(q => q.X).ThenBy(q => q.Y).ToList();
            var uniq = new List<V2>();
            foreach (var q in p)
                if (uniq.Count == 0 || (q - uniq[uniq.Count - 1]).Length > 1e-6) uniq.Add(q);
            if (uniq.Count < 3) return uniq;

            var h = new List<V2>();
            foreach (var q in uniq)
            {
                while (h.Count >= 2 && V2.Cross(h[h.Count - 1] - h[h.Count - 2], q - h[h.Count - 2]) <= 1e-9) h.RemoveAt(h.Count - 1);
                h.Add(q);
            }
            int lower = h.Count + 1;
            for (int i = uniq.Count - 2; i >= 0; i--)
            {
                var q = uniq[i];
                while (h.Count >= lower && V2.Cross(h[h.Count - 1] - h[h.Count - 2], q - h[h.Count - 2]) <= 1e-9) h.RemoveAt(h.Count - 1);
                h.Add(q);
            }
            h.RemoveAt(h.Count - 1);
            return h;
        }

        private static bool OnHullEdge(V2 a, V2 b, List<V2> hull, double tol)
        {
            for (int i = 0; i < hull.Count; i++)
            {
                V2 p = hull[i], q = hull[(i + 1) % hull.Count];
                if (DistToSegment(a, p, q) <= tol && DistToSegment(b, p, q) <= tol) return true;
            }
            return false;
        }

        private static double DistToSegment(V2 x, V2 p, V2 q)
        {
            V2 d = q - p;
            double l2 = d.X * d.X + d.Y * d.Y;
            if (l2 < 1e-12) return (x - p).Length;
            double t = Math.Max(0, Math.Min(1, ((x.X - p.X) * d.X + (x.Y - p.Y) * d.Y) / l2));
            return (x - new V2(p.X + t * d.X, p.Y + t * d.Y)).Length;
        }

        // ---- küçük tipler --------------------------------------------------
        private class Bar
        {
            public string Key, Group, KeyA, KeyB, KeyAxy, KeyBxy;
            public V3 A, B, Mid, Dir;
            public double Length;
            public Bar(string key, FinalMember m)
            {
                Key = key;
                Group = m.group_label;
                A = new V3(m.start_x, m.start_y, m.start_z);
                B = new V3(m.end_x, m.end_y, m.end_z);
                Mid = (A + B) / 2.0;
                V3 d = B - A;
                Length = d.Length;
                Dir = Length > 0 ? d / Length : new V3(0, 0, 0);
                KeyA = NodeKey(A);
                KeyB = NodeKey(B);
                KeyAxy = NodeKeyXY(A);
                KeyBxy = NodeKeyXY(B);
            }
        }

        private class PlaneCand { public V3 N; public double D; public double Tilt; }

        private class PlaneComp
        {
            public PlaneCand Plane;
            public List<Bar> Bars;
            public HashSet<string> KeySet;
            public V3 Centroid;
        }

        private struct V2
        {
            public readonly double X, Y;
            public V2(double x, double y) { X = x; Y = y; }
            public double Length { get { return Math.Sqrt(X * X + Y * Y); } }
            public static V2 operator -(V2 a, V2 b) { return new V2(a.X - b.X, a.Y - b.Y); }
            public static double Cross(V2 a, V2 b) { return a.X * b.Y - a.Y * b.X; }
        }

        private struct V3
        {
            public readonly double X, Y, Z;
            public V3(double x, double y, double z) { X = x; Y = y; Z = z; }
            public double Length { get { return Math.Sqrt(X * X + Y * Y + Z * Z); } }
            public static V3 operator +(V3 a, V3 b) { return new V3(a.X + b.X, a.Y + b.Y, a.Z + b.Z); }
            public static V3 operator -(V3 a, V3 b) { return new V3(a.X - b.X, a.Y - b.Y, a.Z - b.Z); }
            public static V3 operator -(V3 a) { return new V3(-a.X, -a.Y, -a.Z); }
            public static V3 operator *(V3 a, double s) { return new V3(a.X * s, a.Y * s, a.Z * s); }
            public static V3 operator /(V3 a, double s) { return new V3(a.X / s, a.Y / s, a.Z / s); }
            public static double Dot(V3 a, V3 b) { return a.X * b.X + a.Y * b.Y + a.Z * b.Z; }
            public static V3 Cross(V3 a, V3 b) { return new V3(a.Y * b.Z - a.Z * b.Y, a.Z * b.X - a.X * b.Z, a.X * b.Y - a.Y * b.X); }
        }
    }
}
