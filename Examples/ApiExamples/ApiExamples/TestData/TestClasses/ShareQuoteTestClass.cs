namespace Aspose.Words.ApiExamples.HelperClasses.TestClasses
{
    public class ShareQuoteTestClass
    {
        internal ShareQuoteTestClass(int date, int volume, double open, double high, double low, double close)
        {
            Date = date;
            Volume = volume;
            Open = open;
            High = high;
            Low = low;
            Close = close;
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
