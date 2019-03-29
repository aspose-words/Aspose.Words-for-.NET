using System.IO;
using Aspose.Words;
using System;
using Aspose.Words.Fields;
using System.Drawing;
using Aspose.BarCode;
using Aspose.BarCode.Generation;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{

    class GenerateACustomBarCodeImage
    {
        public static void Run()
        {
            // ExStart:GenerateACustomBarCodeImage
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            Document doc = new Document(dataDir + @"GenerateACustomBarCodeImage.docx");

            // Set custom barcode generator
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();
            doc.Save(dataDir + @"GenerateACustomBarCodeImage_out.pdf");
            // ExEnd:GenerateACustomBarCodeImage
        }
    }

    // ExStart:GenerateACustomBarCodeImage_IBarcodeGenerator
    public class CustomBarcodeGenerator : IBarcodeGenerator
    {
        /// <summary>
        /// Converts barcode image height from Word units to Aspose.BarCode units.
        /// </summary>
        /// <param name="heightInTwipsString"></param>
        /// <returns></returns>
        private static float ConvertSymbolHeight(string heightInTwipsString)
        {
            // Input value is in 1/1440 inches (twips)
            int heightInTwips = int.MinValue;
            int.TryParse(heightInTwipsString, out heightInTwips);

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
            int color = int.MinValue;
            int.TryParse(inputColor.Replace("0x", ""), out color);

            if (color == int.MinValue)
                throw new Exception("Error! Incorrect color - " + inputColor + ".");

            return Color.FromArgb(color >> 16, (color & 0xFF00) >> 8, color & 0xFF);

            // Backword conversion -
            //return string.Format("0x{0,6:X6}", mControl.ForeColor.ToArgb() & 0xFFFFFF);
        }

        /// <summary>
        /// Converts bar code scaling factor from percents to float.
        /// </summary>
        /// <param name="scalingFactor"></param>
        /// <returns></returns>
        private static float ConvertScalingFactor(string scalingFactor)
        {
            bool isParsed = false;
            int percents = int.MinValue;
            int.TryParse(scalingFactor, out percents);

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
        /// Implementation of the GetBarCodeImage() method for IBarCodeGenerator interface.
        /// </summary>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public Image GetBarcodeImage(BarcodeParameters parameters)
        {
            if (parameters.BarcodeType == null || parameters.BarcodeValue == null)
                return null;

            string type = parameters.BarcodeType.ToUpper();
            var encodeType = EncodeTypes.None;

            switch (type)
            {
                case "QR":
                    encodeType = Aspose.BarCode.Generation.EncodeTypes.QR;
                    break;
                case "CODE128":
                    encodeType = Aspose.BarCode.Generation.EncodeTypes.Code128;
                    break;
                case "CODE39":
                    encodeType = Aspose.BarCode.Generation.EncodeTypes.Code39Standard;
                    break;
                case "EAN8":
                    encodeType = Aspose.BarCode.Generation.EncodeTypes.EAN8;
                    break;
                case "EAN13":
                    encodeType = Aspose.BarCode.Generation.EncodeTypes.EAN13;
                    break;
                case "UPCA":
                    encodeType = Aspose.BarCode.Generation.EncodeTypes.UPCA;
                    break;
                case "UPCE":
                    encodeType = Aspose.BarCode.Generation.EncodeTypes.UPCE;
                    break;
                case "ITF14":
                    encodeType = Aspose.BarCode.Generation.EncodeTypes.ITF14;
                    break;
                case "CASE":
                    encodeType = Aspose.BarCode.Generation.EncodeTypes.None;
                    break;
            }

            //builder.EncodeType = ConvertBarcodeType(parameters.BarcodeType);
            if (encodeType == Aspose.BarCode.Generation.EncodeTypes.None)
                return null;

            BarCodeGenerator builder = new BarCodeGenerator(encodeType);
            builder.CodeText = parameters.BarcodeValue;

            if (encodeType == Aspose.BarCode.Generation.EncodeTypes.QR)
                builder.D2.DisplayText = parameters.BarcodeValue;

            if (parameters.ForegroundColor != null)
                builder.ForeColor = ConvertColor(parameters.ForegroundColor);

            if (parameters.BackgroundColor != null)
                builder.BackColor = ConvertColor(parameters.BackgroundColor);

            if (parameters.SymbolHeight != null)
            {
                builder.BarCodeHeight.Millimeters = ConvertSymbolHeight(parameters.SymbolHeight);
                builder.AutoSizeMode = AutoSizeMode.Nearest;
            }

            builder.CodeTextStyle.Location = CodeLocation.None;

            if (parameters.DisplayText)
                builder.CodeTextStyle.Location = CodeLocation.Below;

            builder.CaptionAbove.Text = "";

            const float scale = 0.4f; // Empiric scaling factor for converting Word barcode to Aspose.BarCode
            float xdim = 1.0f;

            if (encodeType == Aspose.BarCode.Generation.EncodeTypes.QR)
            {
                builder.AutoSizeMode = AutoSizeMode.Nearest;
                builder.BarCodeWidth.Millimeters *= scale;
                builder.BarCodeHeight.Millimeters = builder.BarCodeWidth.Millimeters;
                xdim = builder.BarCodeHeight.Millimeters / 25;
                builder.XDimension.Millimeters = builder.BarHeight.Millimeters = xdim;
            }

            if (parameters.ScalingFactor != null)
            {
                float scalingFactor = ConvertScalingFactor(parameters.ScalingFactor);
                builder.BarCodeHeight.Millimeters *= scalingFactor;
                if (encodeType == Aspose.BarCode.Generation.EncodeTypes.QR)
                {
                    builder.BarCodeWidth.Millimeters = builder.BarCodeHeight.Millimeters;
                    builder.XDimension.Millimeters = builder.BarHeight.Millimeters = xdim * scalingFactor;
                }

                builder.AutoSizeMode = AutoSizeMode.Nearest;
            }
            return builder.GenerateBarCodeImage();
        }

        //Image IBarcodeGenerator.GetBarcodeImage(BarcodeParameters parameters)
        //{
        //    throw new NotImplementedException();
        //}

        public Image GetOldBarcodeImage(BarcodeParameters parameters)
        {
            throw new NotImplementedException();
        }
    }
    // ExEnd:GenerateACustomBarCodeImage_IBarcodeGenerator
}
