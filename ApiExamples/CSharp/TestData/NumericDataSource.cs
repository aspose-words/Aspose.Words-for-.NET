using System;

namespace ApiExamples.TestData
{
    public class NumericDataSource
    {
        public NumericDataSource(int value1, double value2, int value3, DateTime date)
        {
            Value1 = value1;
            Value2 = value2;
            Value3 = value3;
            Date = date;
        }

        public int Value1 { get; set; }

        public double Value2 { get; set; }

        public int Value3 { get; set; }

        public DateTime Date { get; set; }
    }
}