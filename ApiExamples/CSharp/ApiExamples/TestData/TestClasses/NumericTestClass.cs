using System;

namespace ApiExamples.TestData.TestClasses
{
    public class NumericTestClass
    {
        public int? Value1 { get; set; }
        public double Value2 { get; set; }
        public int Value3 { get; set; }
        public int? Value4 { get; set; }
        public bool Logical { get; set; }
        public DateTime Date { get; set; }

        public NumericTestClass(int? value1, double value2, int value3, int? value4, bool logical, DateTime dateTime)
        {
            this.Value1 = value1;
            this.Value2 = value2;
            this.Value3 = value3;
            this.Value4 = value4;
            this.Logical = logical;
            this.Date = dateTime;
        }

        public int Sum(int value1, int value2)
        {
            int result = value1 + value2;
            return result;
        }
    }
}
