using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.ApiExamples.HelperClasses.TestClasses
{
    public class ShareQuoteTestClass
    {
        internal ShareQuoteTestClass(int date, int volume, double open, double high, double low, double close)
        {
            this.Date = date;
            this.Volume = volume;
            this.Open = open;
            this.High = high;
            this.Low = low;
            this.Close = close;
        }

        public string Color()
        {
            return (Open < Close) ? "#1B9629" : "#96002C";
        }

        public int Date;
        public int Volume;
        public double Open;
        public double High;
        public double Low;
        public double Close;
    }
}
