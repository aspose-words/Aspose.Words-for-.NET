using System.Drawing;

namespace DocsExamples.LINQ_Reporting_Engine.Helpers.Data_Source_Objects
{
    //ExStart:Color
    public class BackgroundColor
    {
        public string Name { get; set; }
        public Color Color { get; set; }
        public int? ColorCode { get; set; } = null;
        public double? Value1 { get; set; } = null;
        public double? Value2 { get; set; } = null;
        public double? Value3 { get; set; } = null;
    }
    //ExEnd:Color
}