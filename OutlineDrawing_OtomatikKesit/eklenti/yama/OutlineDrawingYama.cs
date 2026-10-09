// Tek DLL'de OutlineDrawing'in iki hatasını IL düzeyinde düzeltir (Mono.Cecil):
//  1) dataprocessing.FinalizeMember: sonuna (Z normalizasyonundan önce) SimetriDuzeltici.Uygula(this)
//     -> simetri kodu 12 olan elemanların X aynaları
//  2) dataprocessing.SectionDetection: section sözlüğü anahtarı eleman adı yerine "kesit§eleman"
//     (KesitAnahtari.Olustur) -> bir eleman birden fazla kesitte çizilebilir
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
