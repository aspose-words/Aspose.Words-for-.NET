﻿using System.Drawing;
using ApiExamples.TestData.TestClasses;

namespace ApiExamples.TestData.TestBuilders
{
    public class ColorItemTestBuilder
    {
        public string Name;
        public Color Color;
        public int ColorCode;
        public double Value1;
        public double Value2;
        public double Value3;

        public ColorItemTestBuilder()
        {
            Name = "DefaultName";
            Color = System.Drawing.Color.Black;
            ColorCode = System.Drawing.Color.Black.ToArgb();
            Value1 = 1.0;
            Value2 = 1.0;
            Value3 = 1.0;
        }

        public ColorItemTestBuilder WithColor(string name, Color color)
        {
            this.Name = name;
            this.Color = color;
            return this;
        }

        public ColorItemTestBuilder WithColorCode(string name, int colorCode)
        {
            this.Name = name;
            this.ColorCode = colorCode;
            return this;
        }

        public ColorItemTestBuilder WithColorAndValues(string name, Color color, double value1, double value2, double value3)
        {
            this.Name = name;
            this.Color = color;
            this.Value1 = value1;
            this.Value2 = value2;
            this.Value3 = value3;
            return this;
        }

        public ColorItemTestBuilder WithColorCodeAndValues(string name, int colorCode, double value1, double value2, double value3)
        {
            this.Name = name;
            this.ColorCode = colorCode;
            this.Value1 = value1;
            this.Value2 = value2;
            this.Value3 = value3;
            return this;
        }

        public ColorItemTestClass Build()
        {
            return new ColorItemTestClass(Name, Color, ColorCode, Value1, Value2, Value3);
        }
    }
}
