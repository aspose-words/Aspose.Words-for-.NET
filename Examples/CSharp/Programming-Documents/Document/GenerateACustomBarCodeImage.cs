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
        public Image GetBarcodeImage(Fields.BarcodeParameters parameters)
        {
            if (parameters.BarcodeType == null || parameters.BarcodeValue == null)
                return null;

            string type = parameters.BarcodeType.ToUpper();
            var encodeType = EncodeTypes.None;

            switch (type)
            {
                case "QR":
                    encodeType = EncodeTypes.QR;
                    break;
                case "CODE128":
                    encodeType = EncodeTypes.Code128;
                    break;
                case "CODE39":
                    encodeType = EncodeTypes.Code39Standard;
                    break;
                case "EAN8":
                    encodeType = EncodeTypes.EAN8;
                    break;
                case "EAN13":
                    encodeType = EncodeTypes.EAN13;
                    break;
                case "UPCA":
                    encodeType = EncodeTypes.UPCA;
                    break;
                case "UPCE":
                    encodeType = EncodeTypes.UPCE;
                    break;
                case "ITF14":
                    encodeType = EncodeTypes.ITF14;
                    break;
                case "CASE":
                    encodeType = EncodeTypes.None;
                    break;
            }

            if (encodeType == EncodeTypes.None)
                return null;

            BarcodeGenerator generator = new BarcodeGenerator(encodeType);
            generator.CodeText = parameters.BarcodeValue;

            if (encodeType == EncodeTypes.QR)
                generator.Parameters.Barcode.CodeTextParameters.TwoDDisplayText = parameters.BarcodeValue;

            if (parameters.ForegroundColor != null)
                generator.Parameters.Barcode.ForeColor = ConvertColor(parameters.ForegroundColor);

            if (parameters.BackgroundColor != null)
                generator.Parameters.BackColor = ConvertColor(parameters.BackgroundColor);

            if (parameters.SymbolHeight != null)
            {
                generator.Parameters.Barcode.BarCodeHeight.Millimeters = ConvertSymbolHeight(parameters.SymbolHeight);
                generator.Parameters.Barcode.AutoSizeMode = AutoSizeMode.Nearest;
            }

            generator.Parameters.Barcode.CodeTextParameters.Location = CodeLocation.None;

            if (parameters.DisplayText)
                generator.Parameters.Barcode.CodeTextParameters.Location = CodeLocation.Below;

            generator.Parameters.CaptionAbove.Text = "";

            const float scale = 0.4f; // Empiric scaling factor for converting Word barcode to Aspose.BarCode
            float xdim = 1.0f;

            if (encodeType == EncodeTypes.QR)
            {
                generator.Parameters.Barcode.AutoSizeMode = AutoSizeMode.Nearest;
                generator.Parameters.Barcode.BarCodeWidth.Millimeters *= scale;
                generator.Parameters.Barcode.BarCodeHeight.Millimeters = generator.Parameters.Barcode.BarCodeWidth.Millimeters;
                xdim = generator.Parameters.Barcode.BarCodeHeight.Millimeters / 25;
                generator.Parameters.Barcode.XDimension.Millimeters = generator.Parameters.Barcode.BarHeight.Millimeters = xdim;
            }

            if (parameters.ScalingFactor != null)
            {
                float scalingFactor = ConvertScalingFactor(parameters.ScalingFactor);
                generator.Parameters.Barcode.BarCodeHeight.Millimeters *= scalingFactor;
                if (encodeType == EncodeTypes.QR)
                {
                    generator.Parameters.Barcode.BarCodeWidth.Millimeters = generator.Parameters.Barcode.BarCodeHeight.Millimeters;
                    generator.Parameters.Barcode.XDimension.Millimeters = generator.Parameters.Barcode.BarHeight.Millimeters = xdim * scalingFactor;
                }

                generator.Parameters.Barcode.AutoSizeMode = AutoSizeMode.Nearest;
            }
            return generator.GenerateBarCodeImage();
        }
        
        public Image GetOldBarcodeImage(Fields.BarcodeParameters parameters)
        {
            throw new NotImplementedException();
        }
    }
    // ExEnd:GenerateACustomBarCodeImage_IBarcodeGenerator
}
