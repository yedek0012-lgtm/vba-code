// TEST ICIN sahte Acdbmgd: SectionDetection'in kullandigi geometri tipleri (gercek AutoCAD davranisi)
using System; using System.Reflection;
[assembly: AssemblyVersion("24.1.0.0")]
namespace Autodesk.AutoCAD.Runtime {
  public abstract class DisposableWrapper {
    public static bool operator ==(DisposableWrapper a, DisposableWrapper b) { return ReferenceEquals(a, b); }
    public static bool operator !=(DisposableWrapper a, DisposableWrapper b) { return !ReferenceEquals(a, b); }
    public override bool Equals(object o) { return ReferenceEquals(this, o); }
    public override int GetHashCode() { return base.GetHashCode(); }
  }
}
namespace Autodesk.AutoCAD.Geometry {
  public struct Vector3d {
    double x, y, z;
    public Vector3d(double x, double y, double z) { this.x = x; this.y = y; this.z = z; }
    public double X { get { return x; } } public double Y { get { return y; } } public double Z { get { return z; } }
    public double Length { get { return Math.Sqrt(x*x + y*y + z*z); } }
    public static Vector3d XAxis { get { return new Vector3d(1,0,0); } }
    public static Vector3d YAxis { get { return new Vector3d(0,1,0); } }
    public static Vector3d ZAxis { get { return new Vector3d(0,0,1); } }
    public Vector3d GetNormal() { double l = Length; return l < 1e-300 ? this : new Vector3d(x/l, y/l, z/l); }
    public Vector3d CrossProduct(Vector3d o) { return new Vector3d(y*o.z - z*o.y, z*o.x - x*o.z, x*o.y - y*o.x); }
    public double DotProduct(Vector3d o) { return x*o.x + y*o.y + z*o.z; }
    public Vector3d MultiplyBy(double s) { return new Vector3d(x*s, y*s, z*s); }
    public Vector3d GetPerpendicularVector() { Vector3d a = Math.Abs(x) < 0.015625 && Math.Abs(y) < 0.015625 ? YAxis : ZAxis; return a.CrossProduct(this).GetNormal(); }
    public static Vector3d operator *(Vector3d v, double s) { return v.MultiplyBy(s); }
    public static Vector3d operator -(Vector3d a, Vector3d b) { return new Vector3d(a.x-b.x, a.y-b.y, a.z-b.z); }
    public static Vector3d operator -(Vector3d a) { return new Vector3d(-a.x, -a.y, -a.z); }
  }
  public struct Point3d {
    double x, y, z;
    public Point3d(double x, double y, double z) { this.x = x; this.y = y; this.z = z; }
    public double X { get { return x; } } public double Y { get { return y; } } public double Z { get { return z; } }
    public static Point3d Origin { get { return new Point3d(0,0,0); } }
    public double DistanceTo(Point3d o) { return (this - o).Length; }
    public static Vector3d operator -(Point3d a, Point3d b) { return new Vector3d(a.x-b.x, a.y-b.y, a.z-b.z); }
    public static Point3d operator +(Point3d a, Vector3d v) { return new Point3d(a.x+v.X, a.y+v.Y, a.z+v.Z); }
    public static Point3d operator -(Point3d a, Vector3d v) { return new Point3d(a.x-v.X, a.y-v.Y, a.z-v.Z); }
  }
  public struct Point2d {
    double x, y;
    public Point2d(double x, double y) { this.x = x; this.y = y; }
    public double X { get { return x; } } public double Y { get { return y; } }
  }
  public abstract class PlanarEntity : Autodesk.AutoCAD.Runtime.DisposableWrapper {
    protected Point3d o; protected Vector3d n;
    public virtual Vector3d Normal { get { return n; } }
    public virtual Point3d PointOnPlane { get { return o; } }
  }
  public class Plane : PlanarEntity {
    public Plane(Point3d origin, Vector3d normal) { o = origin; n = normal.GetNormal(); }
  }
}
