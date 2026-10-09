// Tek DLL'de OutlineDrawing.dataprocessing.FinalizeMember'ın sonuna, Z normalizasyonundan önce
// OtomatikKesit.SimetriDuzeltici.Uygula(this) çağrısını ekler (Mono.Cecil).
//   mono FinalizeMemberYama.exe girdi.dll cikti.dll [arama klasörleri...]
using System;
using System.Linq;
using Mono.Cecil;
using Mono.Cecil.Cil;

public static class FinalizeMemberYama
{
    public static int Main(string[] a)
    {
        var resolver = new DefaultAssemblyResolver();
        foreach (string d in a.Skip(2)) resolver.AddSearchDirectory(d);
        var asm = AssemblyDefinition.ReadAssembly(a[0], new ReaderParameters { AssemblyResolver = resolver });
        ModuleDefinition mod = asm.MainModule;

        TypeDefinition dp = mod.GetType("OutlineDrawing.dataprocessing");
        MethodDefinition fin = dp.Methods.Single(m => m.Name == "FinalizeMember" && m.Parameters.Count == 0);
        MethodDefinition duz = mod.GetType("OtomatikKesit.SimetriDuzeltici").Methods.Single(m => m.Name == "Uygula");

        if (fin.Body.Instructions.Any(i => i.OpCode == OpCodes.Call && ((MethodReference)i.Operand).Name == "Uygula"))
        {
            Console.WriteLine("Zaten yamalı.");
            asm.Write(a[1]);
            return 0;
        }

        Instruction norm = fin.Body.Instructions.Single(i => i.OpCode == OpCodes.Call &&
            ((MethodReference)i.Operand).Name == "NormalizeFinalMembersZMinToZero");
        Instruction ld = norm.Previous;
        if (ld.OpCode != OpCodes.Ldarg_0) throw new Exception("Beklenmeyen IL: NormalizeFinalMembersZMinToZero öncesi ldarg.0 değil");

        // ld (ldarg.0, dallanma ve finally sonu hedefi) aynı kalır:
        //   ldarg.0 ; call SimetriDuzeltici.Uygula(object) ; ldarg.0 ; call NormalizeFinalMembersZMinToZero ; ret
        ILProcessor il = fin.Body.GetILProcessor();
        Instruction cagri = il.Create(OpCodes.Call, mod.ImportReference(duz));
        il.InsertAfter(ld, cagri);
        il.InsertAfter(cagri, il.Create(OpCodes.Ldarg_0));

        asm.Write(a[1]);
        Console.WriteLine("FinalizeMember yamalandı: " + a[1]);
        return 0;
    }
}
