// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Drawing;
using System.Globalization;
using Aspose.BarCode;
using Aspose.Words.Fields;
#if NETCOREAPP2_1 || __MOBILE__
using Image = SkiaSharp.SKBitmap;
#endif

namespace ApiExamples
{
    /// <summary>
    /// Sample of custom barcode generator implementation (with underlying Aspose.BarCode module)
    /// </summary>
    public class CustomBarcodeGenerator : ApiExampleBase, IBarcodeGenerator
    {
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
            return (float) (heightInTwips * 25.4 / 1440);
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

            // Backward conversion -
            //return string.Format("0x{0,6:X6}", mControl.ForeColor.ToArgb() & 0xFFFFFF);
        }

        /// <summary>
        /// Converts bar code scaling factor from percent to float.
        /// </summary>
        /// <param name="scalingFactor"></param>
        /// <returns></returns>
        private static float ConvertScalingFactor(string scalingFactor)
        {
            bool isParsed = false;
            int percent = TryParseInt(scalingFactor);

            if (percent != int.MinValue && percent >= 10 && percent <= 10000)
                isParsed = true;

            if (!isParsed)
                throw new Exception("Error! Incorrect scaling factor - " + scalingFactor + ".");

            return percent / 100.0f;
        }

        /// <summary>
        /// Implementation of the GetBarCodeImage() method for IBarCodeGenerator interface.
        /// </summary>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public Image GetBarcodeImage(BarcodeParameters parameters)
        {
            if (parameters.BarcodeType == null || parameters.BarcodeValue == null)
                return null;

            BarCodeBuilder builder = new BarCodeBuilder();

            string type = parameters.BarcodeType.ToUpper();

            switch (type)
            {
                case "QR":
                    builder.EncodeType = Aspose.BarCode.Generation.EncodeTypes.QR;
                    break;
                case "CODE128":
                    builder.EncodeType = Aspose.BarCode.Generation.EncodeTypes.Code128;
                    break;
                case "CODE39":
                    builder.EncodeType = Aspose.BarCode.Generation.EncodeTypes.Code39Standard;
                    break;
                case "EAN8":
                    builder.EncodeType = Aspose.BarCode.Generation.EncodeTypes.EAN8;
                    break;
                case "EAN13":
                    builder.EncodeType = Aspose.BarCode.Generation.EncodeTypes.EAN13;
                    break;
                case "UPCA":
                    builder.EncodeType = Aspose.BarCode.Generation.EncodeTypes.UPCA;
                    break;
                case "UPCE":
                    builder.EncodeType = Aspose.BarCode.Generation.EncodeTypes.UPCE;
                    break;
                case "ITF14":
                    builder.EncodeType = Aspose.BarCode.Generation.EncodeTypes.ITF14;
                    break;
                case "CASE":
                    builder.EncodeType = Aspose.BarCode.Generation.EncodeTypes.None;
                    break;
            }

            if (builder.EncodeType.Equals(Aspose.BarCode.Generation.EncodeTypes.None))
                return null;

            builder.CodeText = parameters.BarcodeValue;

            if (builder.EncodeType.Equals(Aspose.BarCode.Generation.EncodeTypes.QR))
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

            if (builder.EncodeType.Equals(Aspose.BarCode.Generation.EncodeTypes.QR))
            {
                builder.AutoSize = false;
                builder.ImageWidth *= scale;
                builder.ImageHeight = builder.ImageWidth;
                xdim = builder.ImageHeight / 25;
                builder.xDimension = builder.yDimension = xdim;
            }

            if (parameters.ScalingFactor != null)
            {
                float scalingFactor = ConvertScalingFactor(parameters.ScalingFactor);
                builder.ImageHeight *= scalingFactor;
                if (builder.EncodeType.Equals(Aspose.BarCode.Generation.EncodeTypes.QR))
                {
                    builder.ImageWidth = builder.ImageHeight;
                    builder.xDimension = builder.yDimension = xdim * scalingFactor;
                }

                builder.AutoSize = false;
            }
            
#if NET462 || JAVA
            return builder.BarCodeImage;            

#elif NETCOREAPP2_1 || __MOBILE__
            builder.BarCodeImage.Save(ArtifactsDir + "GetBarcodeImage.png");
            return Image.Decode(ArtifactsDir + "OldBarcodeImage.png");
#endif
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

            BarCodeBuilder builder = new BarCodeBuilder
            {
                EncodeType = Aspose.BarCode.Generation.EncodeTypes.Postnet,
                CodeText = parameters.PostalAddress
            };

            // Hardcode type for old-fashioned Barcode
#if NET462 || JAVA
            return builder.BarCodeImage;
#elif NETCOREAPP2_1 || __MOBILE__
            builder.BarCodeImage.Save(ApiExampleBase.ArtifactsDir + "OldBarcodeImage.png");            
            return Image.Decode(ApiExampleBase.ArtifactsDir + "OldBarcodeImage.png");            
#endif
        }

        /// <summary>
        /// Parses an integer using the invariant culture. Returns Int.MinValue if cannot parse.
        /// 
        /// Allows leading sign.
        /// Allows leading and trailing spaces.
        /// </summary>
        public static int TryParseInt(string s)
        {
            return double.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out double temp)
                ? CastDoubleToInt(temp)
                : int.MinValue;
        }

        /// <summary>
        /// Casts a double to int32 in a way that uint32 are "correctly" casted too (they become negative numbers).
        /// </summary>
        public static int CastDoubleToInt(double value)
        {
            long temp = (long) value;
            return (int) temp;
        }

        /// <summary>
        /// Try parses a hex String into an integer value.
        /// on error return int.MinValue
        /// </summary>
        public static int TryParseHex(string s)
        {
            return int.TryParse(s, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int result)
                ? result
                : int.MinValue;
        }
    }
}