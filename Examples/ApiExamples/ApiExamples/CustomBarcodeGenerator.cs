// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
#if NET5_0_OR_GREATER
using SkiaSharp;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
#else
using System.Drawing;
using System.Drawing.Imaging;
#endif

namespace ApiExamples
{
    internal class CustomBarcodeGeneratorUtils
    {
        /// <summary>
        /// Converts a height value in twips to pixels using a default DPI of 96.
        /// </summary>
        /// <param name="heightInTwips">The height value in twips.</param>
        /// <param name="defVal">The default value to return if the conversion fails.</param>
        /// <returns>The height value in pixels.</returns>
        public static double TwipsToPixels(string heightInTwips, double defVal)
        {
            return TwipsToPixels(heightInTwips, 96, defVal);
        }

        /// <summary>
        /// Converts a height value in twips to pixels based on the given resolution.
        /// </summary>
        /// <param name="heightInTwips">The height value in twips to be converted.</param>
        /// <param name="resolution">The resolution in pixels per inch.</param>
        /// <param name="defVal">The default value to be returned if the conversion fails.</param>
        /// <returns>The converted height value in pixels.</returns>
        public static double TwipsToPixels(string heightInTwips, double resolution, double defVal)
        {
            try
            {
                int lVal = int.Parse(heightInTwips);
                return (lVal / 1440.0) * resolution;
            }
            catch
            {
                return defVal;
            }
        }

        /// <summary>
        /// Gets the rotation angle in degrees based on the given rotation angle string.
        /// </summary>
        /// <param name="rotationAngle">The rotation angle string.</param>
        /// <param name="defVal">The default value to return if the rotation angle is not recognized.</param>
        /// <returns>The rotation angle in degrees.</returns>
        public static float GetRotationAngle(string rotationAngle, float defVal)
        {
            switch (rotationAngle)
            {
                case "0":
                    return 0;
                case "1":
                    return 270;
                case "2":
                    return 180;
                case "3":
                    return 90;
                default:
                    return defVal;
            }
        }

        /// <summary>
        /// Converts a string representation of an error correction level to a QRErrorLevel enum value.
        /// </summary>
        /// <param name="errorCorrectionLevel">The string representation of the error correction level.</param>
        /// <param name="def">The default error correction level to return if the input is invalid.</param>
        /// <returns>The corresponding QRErrorLevel enum value.</returns>
        public static QRErrorLevel GetQRCorrectionLevel(string errorCorrectionLevel, QRErrorLevel def)
        {
            switch (errorCorrectionLevel)
            {
                case "0":
                    return QRErrorLevel.LevelL;
                case "1":
                    return QRErrorLevel.LevelM;
                case "2":
                    return QRErrorLevel.LevelQ;
                case "3":
                    return QRErrorLevel.LevelH;
                default:
                    return def;
            }
        }

        /// <summary>
        /// Gets the barcode encode type based on the given encode type from Word.
        /// </summary>
        /// <param name="encodeTypeFromWord">The encode type from Word.</param>
        /// <returns>The barcode encode type.</returns>
        public static SymbologyEncodeType GetBarcodeEncodeType(string encodeTypeFromWord)
        {
            // https://support.microsoft.com/en-au/office/field-codes-displaybarcode-6d81eade-762d-4b44-ae81-f9d3d9e07be3
            switch (encodeTypeFromWord)
            {
                case "QR":
                    return EncodeTypes.QR;
                case "CODE128":
                    return EncodeTypes.Code128;
                case "CODE39":
                    return EncodeTypes.Code39;
                case "JPPOST":
                    return EncodeTypes.RM4SCC;
                case "EAN8":
                case "JAN8":
                    return EncodeTypes.EAN8;
                case "EAN13":
                case "JAN13":
                    return EncodeTypes.EAN13;
                case "UPCA":
                    return EncodeTypes.UPCA;
                case "UPCE":
                    return EncodeTypes.UPCE;
                case "CASE":
                case "ITF14":
                    return EncodeTypes.ITF14;
                case "NW7":
                    return EncodeTypes.Codabar;
                default:
                    return EncodeTypes.None;
            }
        }

        /// <summary>
        /// Converts a hexadecimal color string to a Color object.
        /// </summary>
        /// <param name="inputColor">The hexadecimal color string to convert.</param>
        /// <param name="defVal">The default Color value to return if the conversion fails.</param>
        /// <returns>The Color object representing the converted color, or the default value if the conversion fails.</returns>
        public static Color ConvertColor(string inputColor, Color defVal)
        {
            if (string.IsNullOrEmpty(inputColor)) return defVal;
            try
            {
                int color = Convert.ToInt32(inputColor, 16);
                // Return Color.FromArgb((color >> 16) & 0xFF, (color >> 8) & 0xFF, color & 0xFF);
                return Color.FromArgb(color & 0xFF, (color >> 8) & 0xFF, (color >> 16) & 0xFF);
            }
            catch
            {
                return defVal;
            }
        }

        /// <summary>
        /// Calculates the scale factor based on the provided string representation.
        /// </summary>
        /// <param name="scaleFactor">The string representation of the scale factor.</param>
        /// <param name="defVal">The default value to return if the scale factor cannot be parsed.</param>
        /// <returns>
        /// The scale factor as a decimal value between 0 and 1, or the default value if the scale factor cannot be parsed.
        /// </returns>
        public static double ScaleFactor(string scaleFactor, double defVal)
        {
            try
            {
                int scale = int.Parse(scaleFactor);
                return scale / 100.0;
            }
            catch
            {
                return defVal;
            }
        }

        /// <summary>
        /// Sets the position code style for a barcode generator.
        /// </summary>
        /// <param name="gen">The barcode generator.</param>
        /// <param name="posCodeStyle">The position code style to set.</param>
        /// <param name="barcodeValue">The barcode value.</param>
        public static void SetPosCodeStyle(BarcodeGenerator gen, string posCodeStyle, string barcodeValue)
        {
            switch (posCodeStyle)
            {
                // STD default and without changes.
                case "SUP2":
                    gen.CodeText = barcodeValue.Substring(0, barcodeValue.Length - 2);
                    gen.Parameters.Barcode.Supplement.SupplementData = barcodeValue.Substring(barcodeValue.Length - 2, 2);
                    break;
                case "SUP5":
                    gen.CodeText = barcodeValue.Substring(0, barcodeValue.Length - 5);
                    gen.Parameters.Barcode.Supplement.SupplementData = barcodeValue.Substring(barcodeValue.Length - 5, 5);
                    break;
                case "CASE":
                    gen.Parameters.Border.Visible = true;
                    gen.Parameters.Border.Color = gen.Parameters.Barcode.BarColor;
                    gen.Parameters.Border.DashStyle = BorderDashStyle.Solid;
                    gen.Parameters.Border.Width.Pixels = gen.Parameters.Barcode.XDimension.Pixels * 5;
                    break;
            }
        }

        public const double DefaultQRXDimensionInPixels = 4.0;
        public const double Default1DXDimensionInPixels = 1.0;

        /// <summary>
        /// Draws an error image with the specified exception message.
        /// </summary>
        /// <param name="error">The exception containing the error message.</param>
        /// <returns>A Bitmap object representing the error image.</returns>
        public static Bitmap DrawErrorImage(Exception error)
        {
            Bitmap bmp = new Bitmap(100, 100);

            using (Graphics grf = Graphics.FromImage(bmp))
#if NET5_0_OR_GREATER
                grf.DrawString(error.Message, new Aspose.Drawing.Font("Microsoft Sans Serif", 8f, FontStyle.Regular), Brushes.Red, new Rectangle(0, 0, bmp.Width, bmp.Height));
#else
                grf.DrawString(error.Message, new System.Drawing.Font("Microsoft Sans Serif", 8f, FontStyle.Regular), Brushes.Red, new Rectangle(0,0, bmp.Width, bmp.Height));
#endif
            return bmp;
        }

#if NET5_0_OR_GREATER
        public static SKBitmap ConvertImageToWord(Bitmap bmp)
        {
            MemoryStream ms = new MemoryStream();
            bmp.Save(ms, ImageFormat.Png);
            ms.Position = 0;

            return SKBitmap.Decode(ms);
        }
#else
        public static Image ConvertImageToWord(Bitmap bmp)
        {
            return bmp;
        }
#endif
    }

    internal class CustomBarcodeGenerator : IBarcodeGenerator
    {
#if NET5_0_OR_GREATER
        public SKBitmap GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
#else
        public Image GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
#endif
        {
            try
            {
                BarcodeGenerator gen = new BarcodeGenerator(CustomBarcodeGeneratorUtils.GetBarcodeEncodeType(parameters.BarcodeType), parameters.BarcodeValue);

                // Set color.
                gen.Parameters.Barcode.BarColor = CustomBarcodeGeneratorUtils.ConvertColor(parameters.ForegroundColor, gen.Parameters.Barcode.BarColor);
                gen.Parameters.BackColor = CustomBarcodeGeneratorUtils.ConvertColor(parameters.BackgroundColor, gen.Parameters.BackColor);

                // Set display or hide text.
                if (!parameters.DisplayText)
                    gen.Parameters.Barcode.CodeTextParameters.Location = CodeLocation.None;
                else
                    gen.Parameters.Barcode.CodeTextParameters.Location = CodeLocation.Below;

                // Set QR Code error correction level.s
                gen.Parameters.Barcode.QR.QrErrorLevel = QRErrorLevel.LevelH;
                if (!string.IsNullOrEmpty(parameters.ErrorCorrectionLevel))
                    gen.Parameters.Barcode.QR.QrErrorLevel = CustomBarcodeGeneratorUtils.GetQRCorrectionLevel(parameters.ErrorCorrectionLevel, gen.Parameters.Barcode.QR.QrErrorLevel);

                // Set rotation angle.
                if (!string.IsNullOrEmpty(parameters.SymbolRotation))
                    gen.Parameters.RotationAngle = CustomBarcodeGeneratorUtils.GetRotationAngle(parameters.SymbolRotation, gen.Parameters.RotationAngle);

                // Set scaling factor.
                double scalingFactor = 1;
                if (!string.IsNullOrEmpty(parameters.ScalingFactor))
                    scalingFactor = CustomBarcodeGeneratorUtils.ScaleFactor(parameters.ScalingFactor, scalingFactor);

                // Set size.
                if (gen.BarcodeType == EncodeTypes.QR)
                    gen.Parameters.Barcode.XDimension.Pixels = (float)Math.Max(1.0, Math.Round(CustomBarcodeGeneratorUtils.DefaultQRXDimensionInPixels * scalingFactor));
                else
                    gen.Parameters.Barcode.XDimension.Pixels = (float)Math.Max(1.0, Math.Round(CustomBarcodeGeneratorUtils.Default1DXDimensionInPixels * scalingFactor));

                //Set height.
                if (!string.IsNullOrEmpty(parameters.SymbolHeight))
                    gen.Parameters.Barcode.BarHeight.Pixels = (float)Math.Max(5.0, Math.Round(CustomBarcodeGeneratorUtils.TwipsToPixels(parameters.SymbolHeight, gen.Parameters.Barcode.BarHeight.Pixels) * scalingFactor));

                // Set style of a Point-of-Sale barcode.
                if (!string.IsNullOrEmpty(parameters.PosCodeStyle))
                    CustomBarcodeGeneratorUtils.SetPosCodeStyle(gen, parameters.PosCodeStyle, parameters.BarcodeValue);

                return CustomBarcodeGeneratorUtils.ConvertImageToWord(gen.GenerateBarCodeImage());
            }
            catch (Exception e)
            {
                return CustomBarcodeGeneratorUtils.ConvertImageToWord(CustomBarcodeGeneratorUtils.DrawErrorImage(e));
            }
        }

#if NET5_0_OR_GREATER
        public SKBitmap GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
#else
        public Image GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
#endif
        {
            throw new NotImplementedException();
        }
    }
}
