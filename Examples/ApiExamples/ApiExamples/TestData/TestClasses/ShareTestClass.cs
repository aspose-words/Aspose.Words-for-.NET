using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Common;

namespace Aspose.Words.ApiExamples.HelperClasses.TestClasses
{
    public class ShareTestClass
    {
        internal ShareTestClass(string sector, string industry, string ticker, double weight, double delta)
        {
            this.Sector = sector;
            this.Industry = industry;
            this.Ticker = ticker;
            this.Weight = weight;
            this.Delta = delta;
        }

        public string Title()
        {
            double percentValue = Delta * 100;
            return string.Format("{0}\r\n{1}%", Ticker, percentValue.ToString());
        }

        public string Color()
        {
            const double fullColorDelta = 0.016d;
            const byte unusedColorChannelValue = 80;

            byte r = unusedColorChannelValue;
            byte g = unusedColorChannelValue;
            byte b = unusedColorChannelValue;

            int value =
                unusedColorChannelValue +
                (int)System.Math.Round(System.Math.Abs(Delta) / fullColorDelta *
                    (byte.MaxValue - unusedColorChannelValue));

            if (value > byte.MaxValue)
                value = byte.MaxValue;

            if (Delta < 0)
                r = (byte)value;
            else
                g = (byte)value;

            return string.Format("#{0:X2}{1:X2}{2:X2}", r, g, b);
        }

        public string IndustryColor()
        {
            if (Industry == "Consumer Electronics")
                return "#1B9629";
            else if (Industry == "Software - Infrastructure")
                return "#6029E3";
            else if (Industry == "Semiconductors")
                return "#E38529";
            else if (Industry == "Internet Content & Information")
                return "#964D05";
            else if (Industry == "Entertainment")
                return "#12E32B";
            else if (Industry == "Internet Retail")
                return "#96002C";
            else if (Industry == "Auto Manufactures")
                return "#1EE3A4";
            else if (Industry == "Credit Services")
                return "#D40B70";
            else
                return "#888888";
        }

        public string Sector;
        public string Industry;
        public string Ticker;
        public double Weight;
        public double Delta;
    }
}
