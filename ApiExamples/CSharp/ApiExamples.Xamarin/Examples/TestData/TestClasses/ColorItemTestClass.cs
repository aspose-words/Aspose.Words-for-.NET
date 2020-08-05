using System.Drawing;

namespace ApiExamples.TestData.TestClasses
{
    public class ColorItemTestClass
    {
        public string Name;
        public Color Color;
        public int ColorCode;
        public double Value1;
        public double Value2;
        public double Value3;

        public ColorItemTestClass(string name, Color color, int colorCode, double value1, double value2, double value3)
        {
            Name = name;
            Color = color;
            ColorCode = colorCode;
            Value1 = value1;
            Value2 = value2;
            Value3 = value3;
        }
    }
}