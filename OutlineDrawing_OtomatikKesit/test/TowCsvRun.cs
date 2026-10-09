// tow_to_csv.py ciktisini okuyup AutoSectionDetector'u calistirir; onerilen kesitleri yazar.
//   mono TowCsvRun.exe kule.csv [planes]
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using OutlineDrawing;

public static class TowCsvRun
{
    public static void Main(string[] args)
    {
        var ci = CultureInfo.InvariantCulture;
        var members = new Dictionary<string, FinalMember>();
        foreach (var line in File.ReadAllLines(args[0]))
        {
            var t = line.Split(',');
            if (t.Length < 10) continue;
            members[t[0]] = new FinalMember
            {
                start_x = double.Parse(t[1], ci), start_y = double.Parse(t[2], ci), start_z = double.Parse(t[3], ci),
                end_x = double.Parse(t[4], ci), end_y = double.Parse(t[5], ci), end_z = double.Parse(t[6], ci),
                group_label = t[7], size = t[8], section_label = t[9]
            };
        }

        var opt = new AutoSectionOptions();
        if (args.Length > 1 && args[1] == "planes") opt.DetectPlanes = true;

        var sw = System.Diagnostics.Stopwatch.StartNew();
        var found = AutoSectionDetector.Detect(members, opt);
        Console.WriteLine("{0} eleman, {1} kesit, {2} ms", members.Count, found.Count, sw.ElapsedMilliseconds);
        int i = 0;
        foreach (var c in found)
        {
            Console.WriteLine("{0}: {1}", (char)('A' + i++), c.Description);
            if (c.Kind == AutoSectionKind.Plane)
                foreach (var p in c.Points) Console.WriteLine("      " + AutoSectionDetector.FormatPoint(p));
        }
    }
}
