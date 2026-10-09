// Sentetik kafes kule uretip AutoSectionDetector'u calistirir.
//   Govde : 36 m, tabanda 8 m -> tepede 2 m, 3 m'lik paneller, tek capraz
//           (karsi yuzler ayna -> "yuzleri kesen sahte egik duzlem" tuzagi)
//   Plan  : Z=12000 kose-merkez, Z=21000 baklava (orta noktalar), Z=30000 X
//   Travers: Z=24000 ve Z=30000'de +X/-X, alt baslik yatay, ust baslik egik
using System;
using System.Collections.Generic;
using System.Diagnostics;
using OutlineDrawing;

public static class SyntheticTowerTest
{
    static readonly Dictionary<string, FinalMember> M = new Dictionary<string, FinalMember>();
    static int _n;

    static void Add(double[] a, double[] b, string grp)
    {
        M["m" + (++_n)] = new FinalMember
        {
            start_x = a[0], start_y = a[1], start_z = a[2],
            end_x = b[0], end_y = b[1], end_z = b[2],
            group_label = grp, size = "L50x5", material = "S235"
        };
    }

    static double W(double z, double top, double panel) { return 4000.0 - 3000.0 * z / top; }
    static double[] P(double x, double y, double z) { return new[] { x, y, z }; }
    static double[] Lerp(double[] a, double[] b, double t) { return P(a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t, a[2] + (b[2] - a[2]) * t); }

    static void Build(double top, double panel, int armSegments)
    {
        M.Clear(); _n = 0;
        int nLev = (int)Math.Round(top / panel);
        int[] sx = { 1, -1, -1, 1 }, sy = { -1, -1, 1, 1 };   // koseler: on-sag, on-sol, arka-sol, arka-sag

        for (int i = 0; i < nLev; i++)
        {
            double z0 = i * panel, z1 = (i + 1) * panel, w0 = W(z0, top, panel), w1 = W(z1, top, panel);
            for (int c = 0; c < 4; c++)
                Add(P(sx[c] * w0, sy[c] * w0, z0), P(sx[c] * w1, sy[c] * w1, z1), "LEG" + i);
            // tek capraz, karsi yuzler ayna
            Add(P(w0, -w0, z0), P(-w1, -w1, z1), "D" + i);   // on
            Add(P(w0, w0, z0), P(-w1, w1, z1), "D" + i);     // arka
            Add(P(w0, -w0, z0), P(w1, w1, z1), "E" + i);     // sag
            Add(P(-w0, -w0, z0), P(-w1, w1, z1), "E" + i);   // sol
        }

        for (int i = 1; i <= nLev; i++)
        {
            double z = i * panel, w = W(z, top, panel);
            bool split = Math.Abs(z - 21000) < 1;
            for (int c = 0; c < 4; c++)
            {
                var a = P(sx[c] * w, sy[c] * w, z);
                var b = P(sx[(c + 1) % 4] * w, sy[(c + 1) % 4] * w, z);
                if (split) { var mid = Lerp(a, b, 0.5); Add(a, mid, "H" + i); Add(mid, b, "H" + i); }
                else Add(a, b, "H" + i);
            }
            if (Math.Abs(z - 12000) < 1 || Math.Abs(z - 30000) < 1)
                for (int c = 0; c < 4; c++) Add(P(sx[c] * w, sy[c] * w, z), P(0, 0, z), "P" + i);
            if (split)
                for (int c = 0; c < 4; c++)
                {
                    var m1 = Lerp(P(sx[c] * w, sy[c] * w, z), P(sx[(c + 1) % 4] * w, sy[(c + 1) % 4] * w, z), 0.5);
                    var m2 = Lerp(P(sx[(c + 1) % 4] * w, sy[(c + 1) % 4] * w, z), P(sx[(c + 2) % 4] * w, sy[(c + 2) % 4] * w, z), 0.5);
                    Add(m1, m2, "P" + i);
                }
        }

        foreach (double zc in new[] { 24000.0, 30000.0 })
            foreach (int s in new[] { 1, -1 })
            {
                double zt = zc + 3000, wc = W(zc, top, panel), wt = W(zt, top, panel), L = 6000;
                var rb = new[] { P(s * wc, -wc, zc), P(s * wc, wc, zc) };
                var rt = new[] { P(s * wt, -wt, zt), P(s * wt, wt, zt) };
                var tb = new[] { P(s * (wc + L), -300, zc), P(s * (wc + L), 300, zc) };
                var tt = new[] { P(s * (wc + L), -300, zc + 600), P(s * (wc + L), 300, zc + 600) };
                string g = "X" + (int)(zc / 1000) + (s > 0 ? "R" : "L");
                for (int k = 0; k < armSegments; k++)
                {
                    double t0 = (double)k / armSegments, t1 = (double)(k + 1) / armSegments;
                    for (int j = 0; j < 2; j++)
                    {
                        Add(Lerp(rb[j], tb[j], t0), Lerp(rb[j], tb[j], t1), g + "BC");      // alt baslik
                        Add(Lerp(rt[j], tt[j], t0), Lerp(rt[j], tt[j], t1), g + "TC");      // ust baslik
                        Add(Lerp(rb[j], tb[j], t1), Lerp(rt[j], tt[j], t1), g + "V");       // yan dikme
                        Add(Lerp(rb[j], tb[j], t0), Lerp(rt[j], tt[j], t1), g + "SD");      // yan capraz
                    }
                    Add(Lerp(rb[0], tb[0], t1), Lerp(rb[1], tb[1], t1), g + "BT");          // alt enine
                    Add(Lerp(rt[0], tt[0], t1), Lerp(rt[1], tt[1], t1), g + "TT");          // ust enine
                    Add(Lerp(rb[0], tb[0], t0), Lerp(rb[1], tb[1], t1), g + "BD");          // alt plan capraz
                    Add(Lerp(rt[0], tt[0], t0), Lerp(rt[1], tt[1], t1), g + "TD");          // ust plan capraz
                }
            }
    }

    public static int Main()
    {
        int fail = 0;

        Build(36000, 3000, 3);
        var sw = Stopwatch.StartNew();
        var res = AutoSectionDetector.Detect(M);
        sw.Stop();
        Console.WriteLine("Uye sayisi: {0}, sure: {1} ms", M.Count, sw.ElapsedMilliseconds);
        char letter = 'A';
        foreach (var c in res)
        {
            Console.WriteLine("{0}: {1}", letter++, c.Description);
            if (c.Kind == AutoSectionKind.Plane)
                foreach (var p in c.Points) Console.WriteLine("      " + AutoSectionDetector.FormatPoint(p));
        }

        var heights = new List<double>();
        int planes = 0;
        foreach (var c in res) { if (c.Kind == AutoSectionKind.Height) heights.Add(c.Height.Value); else planes++; }
        double[] expected = { 30000, 24000, 21000, 12000 };
        if (heights.Count != expected.Length) { Console.WriteLine("HATA: yatay kesit sayisi {0}, beklenen {1}", heights.Count, expected.Length); fail++; }
        else for (int i = 0; i < expected.Length; i++)
                if (Math.Abs(heights[i] - expected[i]) > 1) { Console.WriteLine("HATA: kot {0} != {1}", heights[i], expected[i]); fail++; }
        if (planes != 2) { Console.WriteLine("HATA: egik kesit sayisi {0}, beklenen 2 (travers ust duzlemleri, simetri birlesik)", planes); fail++; }
        foreach (var c in res)
            if (c.Kind == AutoSectionKind.Plane && (c.Points.Count != 4 || c.Points[0].X < 0))
            { Console.WriteLine("HATA: egik kesit +X tarafinda 4 nokta olmali"); fail++; }

        // Performans: yogun kule (1 m panel, 6 segmentli travers)
        Build(36000, 1000, 6);
        sw = Stopwatch.StartNew();
        var dense = AutoSectionDetector.Detect(M);
        sw.Stop();
        Console.WriteLine("Yogun kule: {0} uye, {1} kesit, {2} ms", M.Count, dense.Count, sw.ElapsedMilliseconds);
        foreach (var c in dense) Console.WriteLine("   " + c.Description);
        if (dense.Count != 6) { Console.WriteLine("HATA: yogun kulede kesit sayisi {0}, beklenen 6", dense.Count); fail++; }

        Console.WriteLine(fail == 0 ? "TEST BASARILI" : "TEST BASARISIZ (" + fail + ")");
        return fail == 0 ? 0 : 1;
    }
}
