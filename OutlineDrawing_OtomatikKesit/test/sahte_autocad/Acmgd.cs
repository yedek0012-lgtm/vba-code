// TEST ICIN sahte Acmgd
using System; using System.Reflection;
[assembly: AssemblyVersion("24.1.0.0")]
namespace Autodesk.AutoCAD.ApplicationServices {
  public static class Application { public static DocumentCollection DocumentManager { get { return new DocumentCollection(); } } }
}
