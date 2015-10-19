using Aspose.BarCode;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Aspose.BarCodeGenerator.GenerateBarCode
{
    public partial class GenerateBarCode : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string CodeText = Request.QueryString["codetext"];
            string Symbology = Request.QueryString["symbology"];

            if (String.IsNullOrEmpty(CodeText))
                CodeText = "Aspose .NET BarCode Generator for Dynamics CRM";
            
            //Instantiate barcode object
            BarCodeBuilder BarCode = new BarCodeBuilder();

            //Set the Code text for the barcode
            BarCode.CodeText = CodeText;

            //Set the symbology type to Code128
            string SymbologyText = Symbology;
            BarCode.SymbologyType = GetSymbologyType(SymbologyText);
            //Create an instance of resolution and apply on the barcode image with
            //customized resolution settings
            BarCode.Resolution = new Resolution(200f, 400f, ResolutionMode.Customized);
            MemoryStream MemoryStream = new MemoryStream();
            BarCode.Save(MemoryStream, BarCodeImageFormat.Png);


            byte[] byteData = MemoryStream.ToArray();

            Response.Clear();
            Response.ContentType = "image/jpeg";
            Response.BinaryWrite(byteData);
            Response.End();
        }

        private Symbology GetSymbologyType(string SymbologyText)
        {
            switch (SymbologyText)
            {
                case "Codabar":
                    return Symbology.Codabar;
                case "Code11":
                    return Symbology.Code11;
                case "Code39Standard":
                    return Symbology.Code39Standard;
                case "Code39Extended":
                    return Symbology.Code39Extended;
                case "Code93Standard":
                    return Symbology.Code93Standard;
                case "Code93Extended":
                    return Symbology.Code93Extended;
                case "Code128":
                    return Symbology.Code128;
                case "GS1Code128":
                    return Symbology.GS1Code128;
                case "EAN8":
                    return Symbology.EAN8;
                case "EAN13":
                    return Symbology.EAN13;
                case "EAN14":
                    return Symbology.EAN14;
                case "SCC14":
                    return Symbology.SCC14;
                case "SSCC18":
                    return Symbology.SSCC18;
                case "UPCA":
                    return Symbology.UPCA;
                case "UPCE":
                    return Symbology.UPCE;
                case "ISBN":
                    return Symbology.ISBN;
                case "ISSN":
                    return Symbology.ISSN;
                case "ISMN":
                    return Symbology.ISMN;
                case "Standard2of5":
                    return Symbology.Standard2of5;
                case "Interleaved2of5":
                    return Symbology.Interleaved2of5;
                case "Matrix2of5":
                    return Symbology.Matrix2of5;
                case "ItalianPost25":
                    return Symbology.ItalianPost25;
                case "IATA2of5":
                    return Symbology.IATA2of5;
                case "ITF14":
                    return Symbology.ITF14;
                case "ITF6":
                    return Symbology.ITF6;
                case "MSI":
                    return Symbology.MSI;
                case "VIN":
                    return Symbology.VIN;
                case "DeutschePostIdentcode":
                    return Symbology.DeutschePostIdentcode;
                case "DeutschePostLeitcode":
                    return Symbology.DeutschePostLeitcode;
                case "OPC":
                    return Symbology.OPC;
                case "PZN":
                    return Symbology.PZN;
                case "Code16K":
                    return Symbology.Code16K;
                case "Pharmacode":
                    return Symbology.Pharmacode;
                case "DataMatrix":
                    return Symbology.DataMatrix;
                case "QR":
                    return Symbology.QR;
                case "Aztec":
                    return Symbology.Aztec;
                case "Pdf417":
                    return Symbology.Pdf417;
                case "MacroPdf417":
                    return Symbology.MacroPdf417;
                case "AustraliaPost":
                    return Symbology.AustraliaPost;
                case "Postnet":
                    return Symbology.Postnet;
                case "Planet":
                    return Symbology.Planet;
                case "OneCode":
                    return Symbology.OneCode;
                case "RM4SCC":
                    return Symbology.RM4SCC;
                case "DatabarOmniDirectional":
                    return Symbology.DatabarOmniDirectional;
                case "DatabarTruncated":
                    return Symbology.DatabarTruncated;
                case "DatabarLimited":
                    return Symbology.DatabarLimited;
                case "DatabarExpanded":
                    return Symbology.DatabarExpanded;
                case "SingaporePost":
                    return Symbology.SingaporePost;
                case "GS1DataMatrix":
                    return Symbology.GS1DataMatrix;
                case "AustralianPosteParcel":
                    return Symbology.AustralianPosteParcel;
                case "SwissPostParcel":
                    return Symbology.SwissPostParcel;
                case "DatabarExpandedStacked":
                    return Symbology.DatabarExpandedStacked;
                case "DatabarStacked":
                    return Symbology.DatabarStacked;
                case "DatabarStackedOmniDirectional":
                    return Symbology.DatabarStackedOmniDirectional;
                default:
                    return Symbology.Code128;
            }
        }
    }
}