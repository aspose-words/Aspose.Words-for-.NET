// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Drawing;
using System.Globalization;
using Aspose.BarCode.Generation;
using Aspose.Words.Fields;
using BarcodeParameters = Aspose.Words.Fields.BarcodeParameters;
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

            BarcodeGenerator generator = new BarcodeGenerator(EncodeTypes.QR);

            string type = parameters.BarcodeType.ToUpper();

            switch (type)
            {
                case "QR":
                    generator = new BarcodeGenerator(EncodeTypes.QR);
                    break;
                case "CODE128":
                    generator = new BarcodeGenerator(EncodeTypes.Code128);
                    break;
                case "CODE39":
                    generator = new BarcodeGenerator(EncodeTypes.Code39Standard);
                    break;
                case "EAN8":
                    generator = new BarcodeGenerator(EncodeTypes.EAN8);
                    break;
                case "EAN13":
                    generator = new BarcodeGenerator(EncodeTypes.EAN13);
                    break;
                case "UPCA":
                    generator = new BarcodeGenerator(EncodeTypes.UPCA);
                    break;
                case "UPCE":
                    generator = new BarcodeGenerator(EncodeTypes.UPCE);
                    break;
                case "ITF14":
                    generator = new BarcodeGenerator(EncodeTypes.ITF14);
                    break;
                case "CASE":
                    generator = new BarcodeGenerator(EncodeTypes.None);
                    break;
            }

            if (generator.BarcodeType.Equals(EncodeTypes.None))
                return null;

            generator.CodeText = parameters.BarcodeValue;

            if (generator.BarcodeType.Equals(EncodeTypes.QR))
                generator.Parameters.Barcode.CodeTextParameters.TwoDDisplayText = parameters.BarcodeValue;

            if (parameters.ForegroundColor != null)
                generator.Parameters.Barcode.BarColor = ConvertColor(parameters.ForegroundColor);

            if (parameters.BackgroundColor != null)
                generator.Parameters.BackColor = ConvertColor(parameters.BackgroundColor);

            if (parameters.SymbolHeight != null)
            {
                generator.Parameters.ImageHeight.Pixels = ConvertSymbolHeight(parameters.SymbolHeight);
                generator.Parameters.AutoSizeMode = AutoSizeMode.None;
            }

            generator.Parameters.Barcode.CodeTextParameters.Location = CodeLocation.None;

            if (parameters.DisplayText)
                generator.Parameters.Barcode.CodeTextParameters.Location = CodeLocation.Below;

            generator.Parameters.CaptionAbove.Text = "";

            const float scale = 2.4f; // Empiric scaling factor for converting Word barcode to Aspose.BarCode
            float xdim = 1.0f;

            if (generator.BarcodeType.Equals(EncodeTypes.QR))
            {
                generator.Parameters.AutoSizeMode = AutoSizeMode.Nearest;
                generator.Parameters.ImageWidth.Inches *= scale;
                generator.Parameters.ImageHeight.Inches = generator.Parameters.ImageWidth.Inches;
                xdim = generator.Parameters.ImageHeight.Inches / 25;
                generator.Parameters.Barcode.XDimension.Inches = generator.Parameters.Barcode.BarHeight.Inches = xdim;
            }

            if (parameters.ScalingFactor != null)
            {
                float scalingFactor = ConvertScalingFactor(parameters.ScalingFactor);
                generator.Parameters.ImageHeight.Inches *= scalingFactor;
                
                if (generator.BarcodeType.Equals(EncodeTypes.QR))
                {
                    generator.Parameters.ImageWidth.Inches = generator.Parameters.ImageHeight.Inches;
                    generator.Parameters.Barcode.XDimension.Inches = generator.Parameters.Barcode.BarHeight.Inches = xdim * scalingFactor;
                }

                generator.Parameters.AutoSizeMode = AutoSizeMode.None;
            }

#if NET462 || JAVA
            return generator.GenerateBarCodeImage();            

#elif NETCOREAPP2_1 || __MOBILE__
            generator.GenerateBarCodeImage().Save(ArtifactsDir + "GetBarcodeImage.png");
            return Image.Decode(ArtifactsDir + "GetBarcodeImage.png");
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

            BarcodeGenerator generator = new BarcodeGenerator(EncodeTypes.Postnet)
            {
                CodeText = parameters.PostalAddress
            };

            // Hardcode type for old-fashioned Barcode
#if NET462 || JAVA
            return generator.GenerateBarCodeImage();
#elif NETCOREAPP2_1 || __MOBILE__
            generator.GenerateBarCodeImage().Save(ArtifactsDir + "OldBarcodeImage.png");            
            return Image.Decode(ArtifactsDir + "OldBarcodeImage.png");            
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