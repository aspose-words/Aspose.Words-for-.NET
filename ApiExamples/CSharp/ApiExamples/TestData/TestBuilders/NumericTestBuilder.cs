using System;
using ApiExamples.TestData.TestClasses;

namespace ApiExamples.TestData.TestBuilders
{
    public class NumericTestBuilder
    {
        private int? mValue1;
        private double mValue2;
        private int mValue3;
        private int? mValue4;
        private bool mLogical;
        private DateTime mDate;

        public NumericTestBuilder()
        {
            this.mValue1 = 1;
            this.mValue2 = 1.0;
            this.mValue3 = 1;
            this.mValue4 = 1;
            this.mLogical = false;
            this.mDate = new DateTime(2018,01,01);
        }

        public NumericTestBuilder WithValuesAndDate(int? value1, double value2, int value3, int? value4, DateTime dateTime)
        {
            this.mValue1 = value1;
            this.mValue2 = value2;
            this.mValue3 = value3;
            this.mValue4 = value4;
            this.mDate = dateTime;
            return this;
        }

        public NumericTestBuilder WithValuesAndLogical(int? value1, double value2, int value3, int? value4, bool logical)
        {
            this.mValue1 = value1;
            this.mValue2 = value2;
            this.mValue3 = value3;
            this.mValue4 = value4;
            this.mLogical = logical;
            return this;
        }

        public NumericTestClass Build()
        {
            return new NumericTestClass(mValue1, mValue2, mValue3, mValue4, mLogical, mDate);
        }
    }
}
