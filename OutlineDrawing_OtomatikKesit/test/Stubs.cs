// SADECE TEST ICIN: OutlineDrawing.dll'deki tiplerin AutoCAD'siz kopyalari.
// Gercek projeye EKLEMEYIN (projede bu siniflar zaten var).
namespace OutlineDrawing
{
    public class FinalMember
    {
        public double start_x { get; set; }
        public double start_y { get; set; }
        public double start_z { get; set; }
        public double end_x { get; set; }
        public double end_y { get; set; }
        public double end_z { get; set; }
        public string group_label { get; set; }
        public string size { get; set; }
        public string material { get; set; }
        public string section_label { get; set; }
    }

    public class SectionPoint3D
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        public SectionPoint3D() { }
        public SectionPoint3D(double x, double y, double z) { X = x; Y = y; Z = z; }
    }
}
