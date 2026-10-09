// TEST ICIN sahte accoremgd
using System; using System.Reflection;
[assembly: AssemblyVersion("24.1.0.0")]
namespace Autodesk.AutoCAD.EditorInput { public class Editor { public void WriteMessage(string s) { } } }
namespace Autodesk.AutoCAD.ApplicationServices {
  public class Document { public Autodesk.AutoCAD.EditorInput.Editor Editor { get { return new Autodesk.AutoCAD.EditorInput.Editor(); } } }
  public class DocumentCollection { public Document MdiActiveDocument { get { return new Document(); } } }
}
namespace Autodesk.AutoCAD.ApplicationServices.Core {
  public static class Application { public static Autodesk.AutoCAD.ApplicationServices.DocumentCollection DocumentManager { get { return new Autodesk.AutoCAD.ApplicationServices.DocumentCollection(); } } }
}
