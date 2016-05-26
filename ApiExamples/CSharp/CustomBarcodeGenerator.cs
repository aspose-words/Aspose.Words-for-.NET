using System;
using System.Drawing;
using System.Globalization;
using Aspose.BarCode;
using Aspose.Words.Fields;

namespace ApiExamples
{
    /// <summary>
    /// Sample of custom barcode generator implementation (with underlying Aspose.BarCode module)
    /// </summary>
    public class CustomBarcodeGenerator : IBarcodeGenerator
    {
        /// <summary>
        /// Converts barcode type from Word to Aspose.BarCode.
        /// </summary>
        private static Symbology ConvertBarcodeType(string inputCode)
        {
            if (inputCode == null)
                return (Symbology)int.MinValue;

            string type = inputCode.ToUpper();
            Symbology outputCode = (Symbology)int.MinValue;

            switch (type)
            {
                case "QR":
                    outputCode = Symbology.QR;
                    break;
                case "CODE128":
                    outputCode = Symbology.Code128;
                    break;
                case "CODE39":
                    outputCode = Symbology.Code39Standard;
                    break;
                case "EAN8":
                    outputCode = Symbology.EAN8;
                    break;
                case "EAN13":
                    outputCode = Symbology.EAN13;
                    break;
                case "UPCA":
                    outputCode = Symbology.UPCA;
                    break;
                case "UPCE":
                    outputCode = Symbology.UPCE;
                    break;
                case "ITF14":
                    outputCode = Symbology.ITF14;
                    break;
                case "CASE":
                    break;
            }

            return outputCode;
        }

        /// <summary>
        /// Converts barcode image height from Word units to Aspose.BarCode units.
        /// </summary>
        /// <param name="heightInTwipsString"></param>
        /// <returns></returns>
        private static float ConvertSymbolHeight(string heightInTwipsString)
        {
            // Input value is in 1/1440 inches (twips)
            int heightInTwips = TryParseInt(heightInTwipsString);
            if (heightInTwips == int.MinValue)
                throw new Exception("Error! Incorrect height - " + heightInTwipsString + ".");

            // Convert to mm
            return (float)(heightInTwips * 25.4 / 1440);
        }

        /// <summary>
        /// Converts barcode image color from Word to Aspose.BarCode.
        /// </summary>
        /// <param name="inputColor"></param>
        /// <returns></returns>
        private static Color ConvertColor(string inputColor)
        {
            // Input should be from "0x000000" to "0xFFFFFF"
            int color = TryParseHex(inputColor.Replace("0x", ""));
            if (color == int.MinValue)
                throw new Exception("Error! Incorrect color - " + inputColor + ".");

            return Color.FromArgb(color >> 16, (color & 0xFF00) >> 8, color & 0xFF);
        }

        /// <summary>
        /// Converts bar code scaling factor from percents to float.
        /// </summary>
        /// <param name="scalingFactor"></param>
        /// <returns></returns>
        private static float ConvertScalingFactor(string scalingFactor)
        {
            bool isParsed = false;
            int percents = TryParseInt(scalingFactor);

            if (percents != int.MinValue)
            {
                if (percents >= 10 && percents <= 10000)
                    isParsed = true;
            }

            if (!isParsed)
                throw new Exception("Error! Incorrect scaling factor - " + scalingFactor + ".");

            return percents / 100.0f;
        }

        /// <summary>
        /// Implementation of the GetBarcodeImage() method for IBarCodeGenerator interface.
        /// </summary>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public Image GetBarcodeImage(BarcodeParameters parameters)
        {
            if (parameters.BarcodeType == null || parameters.BarcodeValue == null)
                return null;

            BarCodeBuilder builder = new BarCodeBuilder();

            builder.SymbologyType = ConvertBarcodeType(parameters.BarcodeType);
            if (builder.SymbologyType == (Symbology)int.MinValue)
                return null;

            builder.CodeText = parameters.BarcodeValue;

            if (builder.SymbologyType == Symbology.QR)
                builder.Display2DText = parameters.BarcodeValue;

            if (parameters.ForegroundColor != null)
                builder.ForeColor = ConvertColor(parameters.ForegroundColor);

            if (parameters.BackgroundColor != null)
                builder.BackColor = ConvertColor(parameters.BackgroundColor);

            if (parameters.SymbolHeight != null)
            {
                builder.ImageHeight = ConvertSymbolHeight(parameters.SymbolHeight);
                builder.AutoSize = false;
            }

            builder.CodeLocation = CodeLocation.None;

            if (parameters.DisplayText)
                builder.CodeLocation = CodeLocation.Below;

            builder.CaptionAbove.Text = "";

            const float scale = 0.4f; // Empiric scaling factor for converting Word barcode to Aspose.BarCode
            float xdim = 1.0f;

            if (builder.SymbologyType == Symbology.QR)
            {
                builder.AutoSize = false;
                builder.ImageWidth *= scale;
                builder.ImageHeight = builder.ImageWidth;
                xdim = builder.ImageHeight / 25;
                builder.yDimension = xdim;
                builder.xDimension = xdim;
            }

            if (parameters.ScalingFactor != null)
            {
                float scalingFactor = ConvertScalingFactor(parameters.ScalingFactor);
                builder.ImageHeight *= scalingFactor;
                if (builder.SymbologyType == Symbology.QR)
                {
                    builder.ImageWidth = builder.ImageHeight;
                    builder.yDimension = xdim * scalingFactor;
                    builder.xDimension = builder.yDimension;
                }

                builder.AutoSize = false;
            }

            return builder.BarCodeImage;
        }

        /// <summary>
        /// Implementation of the GetOldBarcodeImage() method for IBarCodeGenerator interface.
        /// </summary>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public Image GetOldBarcodeImage(BarcodeParameters parameters)
        {
            if (parameters.PostalAddress == null)
                return null;

            BarCodeBuilder builder = new BarCodeBuilder();

            // Hardcode type for old-fashioned Barcode
            builder.SymbologyType = Symbology.Postnet;
            builder.CodeText = parameters.PostalAddress;

            return builder.BarCodeImage;
        }

        /// <summary>
        /// Parses an integer using the invariant culture. Returns Int.MinValue if cannot parse.
        /// 
        /// Allows leading sign.
        /// Allows leading and trailing spaces.
        /// </summary>
        public static int TryParseInt(string s)
        {
            double temp;
            return (Double.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out temp)) ? CastDoubleToInt(temp) : int.MinValue;
        }

        /// <summary>
        /// Casts a double to int32 in a way that uint32 are "correctly" casted too (they become negative numbers).
        /// </summary>
        public static int CastDoubleToInt(double value)
        {
            long temp = (long)value;
            return (int)temp;
        }

        /// <summary>
        /// Try parses a hex string into an integer value.
        /// on error return int.MinValue
        /// </summary>
        public static int TryParseHex(string s)
        {
            int result;
            return int.TryParse(s, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out result) ? result : int.MinValue;
        }
    }
}
