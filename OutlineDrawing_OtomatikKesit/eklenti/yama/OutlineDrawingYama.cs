// Tek DLL'de OutlineDrawing'in iki hatasını IL düzeyinde düzeltir (Mono.Cecil):
//  1) dataprocessing.FinalizeMember: sonuna (Z normalizasyonundan önce) SimetriDuzeltici.Uygula(this)
//     -> simetri kodu 12 olan elemanların X aynaları
//  2) dataprocessing.SectionDetection: section sözlüğü anahtarı eleman adı yerine "kesit§eleman"
//     (KesitAnahtari.Olustur) -> bir eleman birden fazla kesitte çizilebilir
//  3) dataprocessing.Parse_GroupLabel: grup satırında tam 9 alan yerine en az 9 alan -> PLS-TOWER'ın daha yeni
//     bir sürümü tabloya sütun eklerse grup tablosu yine okunur (yoksa hiçbir grup okunmuyor, bütün etiketlerde
//     profil / malzeme boş kalıyordu). Tablodan sonraki satırlar (dosya yolları, yorumlu başlık) 9'dan az alanlı.
//  4) dataprocessing.IsPoseLine: düğüm / eleman tablolarında "sonraki kayıt var mı" denetimi adın 15 karakterden kısa
//     ve boşluksuz olmasını istiyordu; 15+ karakterli bir ad (ör. PLS'in ürettiği uzun kesir düğümü adları) tablonun
//     geri kalanını sessizce düşürüyordu. Gövde TowYama.PozSatiri(satır) çağrısıyla değiştirilir: boş değil, ';' yok,
//     tek ad (tırnaklı adda boşluk olabilir). Tablo sonu yine ';' içeren başlık satırından anlaşılır.
//  5) main.runtowerfile: görünüşler arası sabit 50000 ve kesitler arası sabit 30000 yerine kule genişliğine göre
//     TowYama.GorunusAraligi / KesitAraligi (en az 50000 / 30000). 30 m'den geniş kesitler komşusunun üstüne
//     biniyordu (YAA-R1'in kol kotu kesiti 28,7 m; daha geniş traversli kulede üst üste).
//   mono OutlineDrawingYama.exe girdi.dll cikti.dll [arama klasörleri...]
using System;
using System.Collections.Generic;
using System.Linq;
using Mono.Cecil;
using Mono.Cecil.Cil;

public static class OutlineDrawingYama
{
    public static int Main(string[] a)
    {
        var resolver = new DefaultAssemblyResolver();
        foreach (string d in a.Skip(2)) resolver.AddSearchDirectory(d);
        var asm = AssemblyDefinition.ReadAssembly(a[0], new ReaderParameters { AssemblyResolver = resolver });
        ModuleDefinition mod = asm.MainModule;
        TypeDefinition dp = mod.GetType("OutlineDrawing.dataprocessing");

        FinalizeMember(mod, dp);
        SectionDetection(mod, dp);
        GrupTablosu(dp);
        PozSatiri(mod, dp);
        Araliklar(mod);

        asm.Write(a[1]);
        Console.WriteLine("Yamalandı: " + a[1]);
        return 0;
    }

    private static bool Cagri(Instruction i, string ad)
    {
        return (i.OpCode == OpCodes.Call || i.OpCode == OpCodes.Callvirt) && ((MethodReference)i.Operand).Name == ad;
    }

    private static void FinalizeMember(ModuleDefinition mod, TypeDefinition dp)
    {
        MethodDefinition fin = dp.Methods.Single(m => m.Name == "FinalizeMember" && m.Parameters.Count == 0);
        if (fin.Body.Instructions.Any(i => Cagri(i, "Uygula"))) { Console.WriteLine("FinalizeMember zaten yamalı."); return; }
        MethodDefinition duz = mod.GetType("OtomatikKesit.SimetriDuzeltici").Methods.Single(m => m.Name == "Uygula");

        Instruction norm = fin.Body.Instructions.Single(i => Cagri(i, "NormalizeFinalMembersZMinToZero"));
        Instruction ld = norm.Previous;
        if (ld.OpCode != OpCodes.Ldarg_0) throw new Exception("FinalizeMember: beklenmeyen IL");
        // ld (dallanma ve finally sonu hedefi) aynı kalır:
        //   ldarg.0 ; call SimetriDuzeltici.Uygula(object) ; ldarg.0 ; call NormalizeFinalMembersZMinToZero
        ILProcessor il = fin.Body.GetILProcessor();
        Instruction cagri = il.Create(OpCodes.Call, mod.ImportReference(duz));
        il.InsertAfter(ld, cagri);
        il.InsertAfter(cagri, il.Create(OpCodes.Ldarg_0));
        Console.WriteLine("FinalizeMember yamalandı.");
    }

    private static void Araliklar(ModuleDefinition mod)
    {
        TypeDefinition main = mod.Types.FirstOrDefault(t => t.Methods.Any(m => m.Name == "runtowerfile"));
        if (main == null) throw new Exception("runtowerfile bulunamadı");
        MethodDefinition run = main.Methods.Single(m => m.Name == "runtowerfile");
        List<Instruction> ins = run.Body.Instructions.ToList();
        if (ins.Any(i => Cagri(i, "KesitAraligi"))) { Console.WriteLine("runtowerfile zaten yamalı."); return; }
        // dataprocessing.Run(...) sonucunun yerel değişkeni
        Instruction runCagri = ins.First(i => Cagri(i, "Run") && ((MethodReference)i.Operand).DeclaringType.Name == "dataprocessing");
        Instruction st = runCagri.Next;
        VariableDefinition dpYerel = YerelOf(st, run);
        if (dpYerel == null) throw new Exception("runtowerfile: dataprocessing yereli bulunamadı");
        TypeDefinition ty = mod.GetType("OtomatikKesit.TowYama");
        MethodReference gor = mod.ImportReference(ty.Methods.Single(m => m.Name == "GorunusAraligi"));
        MethodReference kes = mod.ImportReference(ty.Methods.Single(m => m.Name == "KesitAraligi"));
        ILProcessor il = run.Body.GetILProcessor();
        int g = 0, k = 0;
        foreach (Instruction i in ins)
        {
            if (i.OpCode != OpCodes.Ldc_I4) continue;
            int v = (int)i.Operand;
            if (v != 50000 && v != 30000) continue;
            // "ldc.i4 N"  ->  "ldloc dp ; call TowYama.XxxAraligi(object)" (dallanma hedefi aynı komut kalır)
            i.OpCode = OpCodes.Ldloc; i.Operand = dpYerel;
            il.InsertAfter(i, il.Create(OpCodes.Call, v == 50000 ? gor : kes));
            if (v == 50000) g++; else k++;
        }
        if (g != 2 || k != 1) throw new Exception("runtowerfile: 2 görünüş + 1 kesit aralığı bekleniyordu, bulunan " + g + " + " + k);
        Console.WriteLine("runtowerfile yamalandı (görünüş / kesit aralığı kule genişliğine göre).");
    }

    private static VariableDefinition YerelOf(Instruction st, MethodDefinition m)
    {
        if (st.OpCode == OpCodes.Stloc_0) return m.Body.Variables[0];
        if (st.OpCode == OpCodes.Stloc_1) return m.Body.Variables[1];
        if (st.OpCode == OpCodes.Stloc_2) return m.Body.Variables[2];
        if (st.OpCode == OpCodes.Stloc_3) return m.Body.Variables[3];
        if (st.OpCode == OpCodes.Stloc_S || st.OpCode == OpCodes.Stloc) return (VariableDefinition)st.Operand;
        return null;
    }

    private static void PozSatiri(ModuleDefinition mod, TypeDefinition dp)
    {
        MethodDefinition ip = dp.Methods.Single(m => m.Name == "IsPoseLine");
        if (ip.Body.Instructions.Any(i => Cagri(i, "PozSatiri"))) { Console.WriteLine("IsPoseLine zaten yamalı."); return; }
        if (ip.Parameters.Count != 1 || ip.ReturnType.FullName != "System.Boolean") throw new Exception("IsPoseLine: beklenmeyen imza");
        MethodDefinition yeni = mod.GetType("OtomatikKesit.TowYama").Methods.Single(m => m.Name == "PozSatiri");
        ip.Body.ExceptionHandlers.Clear();
        ip.Body.Instructions.Clear();
        ILProcessor il = ip.Body.GetILProcessor();
        il.Append(il.Create(ip.IsStatic ? OpCodes.Ldarg_0 : OpCodes.Ldarg_1));
        il.Append(il.Create(OpCodes.Call, mod.ImportReference(yeni)));
        il.Append(il.Create(OpCodes.Ret));
        Console.WriteLine("IsPoseLine yamalandı (uzun / boşluklu ad).");
    }

    private static void GrupTablosu(TypeDefinition dp)
    {
        MethodDefinition pg = dp.Methods.Single(m => m.Name == "Parse_GroupLabel");
        List<Instruction> ins = pg.Body.Instructions.ToList();
        // "ldlen ; conv.i4 ; ldc.i4.s 9 ; ceq ; ldc.i4.0 ; ceq"  (alan sayısı != 9 ise dur)
        for (int i = 0; i + 5 < ins.Count; i++)
        {
            if (ins[i].OpCode != OpCodes.Ldlen || ins[i + 1].OpCode != OpCodes.Conv_I4) continue;
            Instruction dokuz = ins[i + 2];
            bool dokuzMu = (dokuz.OpCode == OpCodes.Ldc_I4_S && Convert.ToInt32(dokuz.Operand) == 9) ||
                           (dokuz.OpCode == OpCodes.Ldc_I4 && Convert.ToInt32(dokuz.Operand) == 9);
            if (!dokuzMu) continue;
            if (ins[i + 3].OpCode == OpCodes.Clt) { Console.WriteLine("Parse_GroupLabel zaten yamalı."); return; }
            if (ins[i + 3].OpCode != OpCodes.Ceq || ins[i + 4].OpCode != OpCodes.Ldc_I4_0 || ins[i + 5].OpCode != OpCodes.Ceq)
                throw new Exception("Parse_GroupLabel: beklenmeyen IL");
            // alan sayısı < 9 ise dur
            ins[i + 3].OpCode = OpCodes.Clt; ins[i + 3].Operand = null;
            ins[i + 4].OpCode = OpCodes.Nop; ins[i + 4].Operand = null;
            ins[i + 5].OpCode = OpCodes.Nop; ins[i + 5].Operand = null;
            Console.WriteLine("Parse_GroupLabel yamalandı (en az 9 alan).");
            return;
        }
        throw new Exception("Parse_GroupLabel: alan sayısı denetimi bulunamadı");
    }

    private static void SectionDetection(ModuleDefinition mod, TypeDefinition dp)
    {
        MethodDefinition sd = dp.Methods.Single(m => m.Name == "SectionDetection");
        if (sd.Body.Instructions.Any(i => Cagri(i, "Olustur"))) { Console.WriteLine("SectionDetection zaten yamalı."); return; }
        MethodDefinition olustur = mod.GetType("OtomatikKesit.KesitAnahtari").Methods.Single(m => m.Name == "Olustur");

        List<Instruction> ins = sd.Body.Instructions.ToList();
        // Kesit satırı (SectionRowData) yerel değişkeni ve get_SectionName: "ldloc V ; callvirt get_SectionName"
        Instruction sn = ins.First(i => Cagri(i, "get_SectionName"));
        Instruction satir = sn.Previous;
        if (!satir.OpCode.Name.StartsWith("ldloc")) throw new Exception("SectionDetection: kesit satırı yereli bulunamadı");

        var hedefler = new List<Instruction>();
        // 1) Yükseklik kesiti: section[tuple.Item1] = new Section{...}  -> "ldfld Item1 ; newobj Section::.ctor"
        hedefler.AddRange(ins.Where(i => i.OpCode == OpCodes.Ldfld && ((FieldReference)i.Operand).Name == "Item1" &&
            i.Next.OpCode == OpCodes.Newobj && ((MethodReference)i.Next.Operand).DeclaringType.Name == "Section"));
        // 2) Nokta kesiti: section[kv.Key] = CreateSectionFromMember(...) -> "call get_section ; ldloca kv ; call get_Key"
        hedefler.AddRange(ins.Where(i => Cagri(i, "get_Key") && i.Previous != null && i.Previous.Previous != null &&
            Cagri(i.Previous.Previous, "get_section")));
        // 3) Nokta kesiti hizalama listesi: hitKeys.Add(kv.Key) -> "call get_Key ; callvirt Add"
        hedefler.AddRange(ins.Where(i => Cagri(i, "get_Key") && i.Next != null && Cagri(i.Next, "Add")));
        if (hedefler.Count != 3) throw new Exception("SectionDetection: 3 anahtar yeri bekleniyordu, bulunan " + hedefler.Count);

        ILProcessor il = sd.Body.GetILProcessor();
        MethodReference olusturRef = mod.ImportReference(olustur);
        foreach (Instruction h in hedefler)
        {
            // yığında: anahtar  ->  anahtar, kesitAdı  ->  Olustur(anahtar, kesitAdı)
            Instruction yukle = satir.Operand != null ? il.Create(satir.OpCode, (VariableDefinition)satir.Operand) : il.Create(satir.OpCode);
            Instruction ad = il.Create(sn.OpCode, (MethodReference)sn.Operand);
            Instruction cagri = il.Create(OpCodes.Call, olusturRef);
            il.InsertAfter(h, yukle);
            il.InsertAfter(yukle, ad);
            il.InsertAfter(ad, cagri);
        }
        Console.WriteLine("SectionDetection yamalandı (3 anahtar yeri).");
    }
}
