// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    internal class ExHtmlSaveOptions : ApiExampleBase
    {
        #region PageMargins

        //For assert this test you need to open html docs and they shouldn't have negative left margins //ToDo: Need to add gold assert tests
        [Test]
        [TestCase(SaveFormat.Html)]
        [TestCase(SaveFormat.Mhtml)]
        [TestCase(SaveFormat.Epub)]
        public void ExportPageMargins(SaveFormat saveFormat)
        {
            var doc = new Document(MyDir + "HtmlSaveOptions.ExportPageMargins.docx");

            var saveOptions = new HtmlSaveOptions();
            saveOptions.SaveFormat = saveFormat;
            saveOptions.ExportPageMargins = true;

            Save(doc, @"\Artifacts\HtmlSaveOptions.ExportPageMargins." + saveFormat.ToString().ToLower(), saveFormat, saveOptions);
        }

        #endregion

        #region HtmlOfficeMathOutputMode

        [Test]
        [TestCase(SaveFormat.Html, HtmlOfficeMathOutputMode.Image)]
        [TestCase(SaveFormat.Mhtml, HtmlOfficeMathOutputMode.MathML)]
        [TestCase(SaveFormat.Epub, HtmlOfficeMathOutputMode.Text)]
        public void ExportOfficeMath(SaveFormat saveFormat, HtmlOfficeMathOutputMode outputMode)
        {
            var doc = new Document(MyDir + "OfficeMath.docx");

            var saveOptions = new HtmlSaveOptions();
            saveOptions.OfficeMathOutputMode = outputMode;

            Save(doc, @"\Artifacts\HtmlSaveOptions.ExportToHtmlUsingImage." + saveFormat.ToString().ToLower(), saveFormat, saveOptions);

            switch (saveFormat)
            {
                case SaveFormat.Html:
                    DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\HtmlSaveOptions.ExportToHtmlUsingImage." + saveFormat.ToString().ToLower(), "<img src=\"HtmlSaveOptions.ExportToHtmlUsingImage.001.png\" width=\"49\" height=\"21\" alt=\"\" style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />");
                    return;

                case SaveFormat.Mhtml:
                    DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\HtmlSaveOptions.ExportToHtmlUsingImage." + saveFormat.ToString().ToLower(), "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>A</mi><mo>=</mo><mi>π</mi><msup><mrow><mi>r</mi></mrow><mrow><mn>2</mn></mrow></msup></math>");
                    return;

                case SaveFormat.Epub:
                    DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\HtmlSaveOptions.ExportToHtmlUsingImage." + saveFormat.ToString().ToLower(), "<span style=\"font-family:\'Cambria Math\'\">A=π</span><span style=\"font-family:\'Cambria Math\'\">r</span><span style=\"font-family:\'Cambria Math\'\">2</span>");
                    return;
            }
        }

        #endregion

        #region ExportTextBoxAsSvg

        [Test]
        [TestCase(SaveFormat.Html, true, Description = "TextBox as svg (html)")]
        [TestCase(SaveFormat.Epub, true, Description = "TextBox as svg (epub)")]
        [TestCase(SaveFormat.Mhtml, false, Description = "TextBox as img (mhtml)")]
        public void ExportTextBoxAsSvg(SaveFormat saveFormat, bool textBoxAsSvg)
        {
            string[] dirFiles;

            var doc = new Document(MyDir + "HtmlSaveOptions.ExportTextBoxAsSvg.docx");

            var saveOptions = new HtmlSaveOptions();
            saveOptions.ExportTextBoxAsSvg = textBoxAsSvg;

            Save(doc, @"\Artifacts\HtmlSaveOptions.ExportTextBoxAsSvg." + saveFormat.ToString().ToLower(), saveFormat, saveOptions);

            switch (saveFormat)
            {
                case SaveFormat.Html:

                    dirFiles = Directory.GetFiles(MyDir + @"\Artifacts\", "HtmlSaveOptions.ExportTextBoxAsSvg.001.png", SearchOption.AllDirectories);
                    Assert.IsEmpty(dirFiles);

                    DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\HtmlSaveOptions.ExportTextBoxAsSvg." + saveFormat.ToString().ToLower(), "﻿<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" width=\"238\" height=\"185\"><defs><clipPath id=\"clip1\"><path d=\"M0,3.600000143 L178.720001221,3.600000143 L178.720001221,84.055038452 L0,84.055038452 Z\" clip-rule=\"evenodd\" /></clipPath></defs><g transform=\"scale(1.33333)\"><g><g><g transform=\"matrix(1,0,0,1,0,0)\"><path d=\"M0,0 L178.720001221,0 L178.720001221,0 L178.720001221,87.655036926 L178.720001221,87.655036926 L0,87.655036926 Z\" fill=\"#ffffff\" fill-rule=\"evenodd\" /><path d=\"M0,0 L178.720001221,0 L178.720001221,0 L178.720001221,87.655036926 L178.720001221,87.655036926 L0,87.655036926 Z\" stroke-width=\"0.75\" stroke-miterlimit=\"10\" stroke=\"#000000\" fill=\"none\" fill-rule=\"evenodd\" /><g transform=\"matrix(1,0,0,1,0,0)\" clip-path=\"url(#clip1)\"><g transform=\"matrix(1,0,0,1,7.200000286,3.600000143)\"><text><tspan x=\"0\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">[Grab</tspan><tspan x=\"25.195999146\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"27.683000565\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">your</tspan><tspan x=\"48.076999664\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"50.564002991\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">reader’s</tspan><tspan x=\"87.275001526\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"89.762001038\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">attention</tspan><tspan x=\"131.442001343\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"133.929000854\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">with</tspan><tspan x=\"153.781005859\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"156.268005371\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">a</tspan><tspan x=\"161.537002563\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"0\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">great</tspan><tspan x=\"23.438999176\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"25.926002502\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">quote</tspan><tspan x=\"52.443004608\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"54.930000305\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">from</tspan><tspan x=\"76.709999084\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"79.196998596\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">the</tspan><tspan x=\"94.134010315\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"96.621002197\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">document</tspan><tspan x=\"142.356002808\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"144.843002319\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">or</tspan><tspan x=\"154.479003906\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"0\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">use</tspan><tspan x=\"15.555000305\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"18.041999817\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">this</tspan><tspan x=\"34.333000183\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"36.819999695\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">space</tspan><tspan x=\"62.295001984\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"64.781997681\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">to</tspan><tspan x=\"74.266998291\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"76.753997803\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">emphasize</tspan><tspan x=\"124.486999512\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"126.973999023\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">a</tspan><tspan x=\"132.242996216\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"134.729995728\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">key</tspan><tspan x=\"150.182998657\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"0\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">point.</tspan><tspan x=\"26.345001221\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"28.832000732\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">To</tspan><tspan x=\"39.993000031\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"42.479999542\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">place</tspan><tspan x=\"66.177001953\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"68.664001465\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">this</tspan><tspan x=\"84.955001831\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"87.442001343\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">text</tspan><tspan x=\"105.047996521\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"107.535003662\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">box</tspan><tspan x=\"123.878997803\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"0\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">anywhere</tspan><tspan x=\"44.451000214\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"46.93800354\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">on</tspan><tspan x=\"58.518001556\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"61.005001068\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">the</tspan><tspan x=\"75.942001343\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"78.429000854\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">page,</tspan><tspan x=\"102.873001099\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"105.36000061\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">just</tspan><tspan x=\"121.758003235\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"124.245002747\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">drag</tspan><tspan x=\"144.305999756\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"146.792999268\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">it.]</tspan></text></g></g></g></g></g></g></svg>");
                    return;

                case SaveFormat.Epub:

                    dirFiles = Directory.GetFiles(MyDir + @"\Artifacts\", "HtmlSaveOptions.ExportTextBoxAsSvg.001.png", SearchOption.AllDirectories);
                    Assert.IsEmpty(dirFiles);

                    DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\HtmlSaveOptions.ExportTextBoxAsSvg." + saveFormat.ToString().ToLower(), "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" width=\"238\" height=\"185\"><defs><clipPath id=\"clip1\"><path d=\"M0,3.600000143 L178.720001221,3.600000143 L178.720001221,84.055038452 L0,84.055038452 Z\" clip-rule=\"evenodd\" /></clipPath></defs><g transform=\"scale(1.33333)\"><g><g><g transform=\"matrix(1,0,0,1,0,0)\"><path d=\"M0,0 L178.720001221,0 L178.720001221,0 L178.720001221,87.655036926 L178.720001221,87.655036926 L0,87.655036926 Z\" fill=\"#ffffff\" fill-rule=\"evenodd\" /><path d=\"M0,0 L178.720001221,0 L178.720001221,0 L178.720001221,87.655036926 L178.720001221,87.655036926 L0,87.655036926 Z\" stroke-width=\"0.75\" stroke-miterlimit=\"10\" stroke=\"#000000\" fill=\"none\" fill-rule=\"evenodd\" /><g transform=\"matrix(1,0,0,1,0,0)\" clip-path=\"url(#clip1)\"><g transform=\"matrix(1,0,0,1,7.200000286,3.600000143)\"><text><tspan x=\"0\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">[Grab</tspan><tspan x=\"25.195999146\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"27.683000565\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">your</tspan><tspan x=\"48.076999664\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"50.564002991\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">reader’s</tspan><tspan x=\"87.275001526\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"89.762001038\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">attention</tspan><tspan x=\"131.442001343\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"133.929000854\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">with</tspan><tspan x=\"153.781005859\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"156.268005371\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">a</tspan><tspan x=\"161.537002563\" y=\"10.473999977\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"0\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">great</tspan><tspan x=\"23.438999176\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"25.926002502\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">quote</tspan><tspan x=\"52.443004608\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"54.930000305\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">from</tspan><tspan x=\"76.709999084\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"79.196998596\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">the</tspan><tspan x=\"94.134010315\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"96.621002197\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">document</tspan><tspan x=\"142.356002808\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"144.843002319\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">or</tspan><tspan x=\"154.479003906\" y=\"24.965000153\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"0\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">use</tspan><tspan x=\"15.555000305\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"18.041999817\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">this</tspan><tspan x=\"34.333000183\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"36.819999695\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">space</tspan><tspan x=\"62.295001984\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"64.781997681\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">to</tspan><tspan x=\"74.266998291\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"76.753997803\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">emphasize</tspan><tspan x=\"124.486999512\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"126.973999023\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">a</tspan><tspan x=\"132.242996216\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"134.729995728\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">key</tspan><tspan x=\"150.182998657\" y=\"39.456001282\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"0\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">point.</tspan><tspan x=\"26.345001221\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"28.832000732\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">To</tspan><tspan x=\"39.993000031\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"42.479999542\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">place</tspan><tspan x=\"66.177001953\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"68.664001465\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">this</tspan><tspan x=\"84.955001831\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"87.442001343\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">text</tspan><tspan x=\"105.047996521\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"107.535003662\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">box</tspan><tspan x=\"123.878997803\" y=\"53.946998596\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"0\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">anywhere</tspan><tspan x=\"44.451000214\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"46.93800354\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">on</tspan><tspan x=\"58.518001556\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"61.005001068\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">the</tspan><tspan x=\"75.942001343\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"78.429000854\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">page,</tspan><tspan x=\"102.873001099\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"105.36000061\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">just</tspan><tspan x=\"121.758003235\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"124.245002747\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">drag</tspan><tspan x=\"144.305999756\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\"> </tspan><tspan x=\"146.792999268\" y=\"68.43800354\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"11\" fill=\"#000000\">it.]</tspan></text></g></g></g></g></g></g></svg>");
                    return;

                case SaveFormat.Mhtml:

                    dirFiles = Directory.GetFiles(MyDir + @"\Artifacts\", "HtmlSaveOptions.ExportTextBoxAsSvg.001.png", SearchOption.AllDirectories);
                    Assert.IsNotEmpty(dirFiles);

                    DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\HtmlSaveOptions.ExportTextBoxAsSvg." + saveFormat.ToString().ToLower(), "<img src=\"HtmlSaveOptions.ExportTextBoxAsSvg.001.png\" width=\"240\" height=\"118\" alt=\"\" style=\"margin:3.22pt 9pt 3.6pt 8.62pt; -aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:14.4pt; -aw-wrap-type:square; float:left\" />");
                    return;
            }
        }

        #endregion

        private static Document Save(Document inputDoc, string outputDocPath, SaveFormat saveFormat, SaveOptions saveOptions)
        {
            switch (saveFormat)
            {
                case SaveFormat.Html:
                    inputDoc.Save(MyDir + outputDocPath, saveOptions);
                    return inputDoc;
                case SaveFormat.Mhtml:
                    inputDoc.Save(MyDir + outputDocPath, saveOptions);
                    return inputDoc;
                case SaveFormat.Epub:
                    inputDoc.Save(MyDir + outputDocPath, saveOptions);
                    //There is draw images bug with epub. Need write to NSezganov
                    return inputDoc;
            }

            return null;
        }

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void ExportUrlForLinkedImage(bool export)
        {
            var doc = new Document(MyDir + "HtmlSaveOptions.ExportUrlForLinkedImage.docx");

            var saveOptions = new HtmlSaveOptions();
            saveOptions.ExportOriginalUrlForLinkedImages = export;

            doc.Save(MyDir + @"\Artifacts\HtmlSaveOptions.ExportUrlForLinkedImage.html", saveOptions);

            var dirFiles = Directory.GetFiles(MyDir + @"\Artifacts\", "HtmlSaveOptions.ExportUrlForLinkedImage.001.png", SearchOption.AllDirectories);

            if (dirFiles.Length == 0)
                DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\HtmlSaveOptions.ExportUrlForLinkedImage.html", "<img src=\"http://www.aspose.com/images/aspose-logo.gif\"");
            else
                DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\HtmlSaveOptions.ExportUrlForLinkedImage.html", "<img src=\"HtmlSaveOptions.ExportUrlForLinkedImage.001.png\"");
        }

        [Ignore("Bug, css styles starting with -aw, even if ExportRoundtripInformation is false")]
        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void ExportRoundtripInformation(bool valueHtml)
        {
            var doc = new Document(MyDir + "HtmlSaveOptions.ExportPageMargins.docx");

            var saveOptions = new HtmlSaveOptions();
            saveOptions.ExportRoundtripInformation = valueHtml;

            doc.Save(MyDir + @"\Artifacts\HtmlSaveOptions.RoundtripInformation.html");

            if (valueHtml)
                DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\HtmlSaveOptions.RoundtripInformation.html", "<img src=\"HtmlSaveOptions.RoundtripInformation.003.png\" width=\"226\" height=\"132\" alt=\"\" style=\"margin-top:-53.74pt; margin-left:-26.75pt; -aw-left-pos:-26.25pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:41.25pt; -aw-wrap-type:none; position:absolute\" /></span><span style=\"height:0pt; display:block; position:absolute; z-index:1\"><img src=\"HtmlSaveOptions.RoundtripInformation.002.png\" width=\"227\" height=\"132\" alt=\"\" style=\"margin-top:74.51pt; margin-left:-23pt; -aw-left-pos:-22.5pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:169.5pt; -aw-wrap-type:none; position:absolute\" /></span><span style=\"height:0pt; display:block; position:absolute; z-index:2\"><img src=\"HtmlSaveOptions.RoundtripInformation.001.png\" width=\"227\" height=\"132\" alt=\"\" style=\"margin-top:199.01pt; margin-left:-23pt; -aw-left-pos:-22.5pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:294pt; -aw-wrap-type:none; position:absolute\" />");
            else
                DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\HtmlSaveOptions.RoundtripInformation.html", "<img src=\"HtmlSaveOptions.RoundtripInformation.003.png\" width=\"226\" height=\"132\" alt=\"\" style=\"margin-top:-53.74pt; margin-left:-26.75pt; -aw-left-pos:-26.25pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:41.25pt; -aw-wrap-type:none; position:absolute\" /></span><span style=\"height:0pt; display:block; position:absolute; z-index:1\"><img src=\"HtmlSaveOptions.RoundtripInformation.002.png\" width=\"227\" height=\"132\" alt=\"\" style=\"margin-top:74.51pt; margin-left:-23pt; -aw-left-pos:-22.5pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:169.5pt; -aw-wrap-type:none; position:absolute\" /></span><span style=\"height:0pt; display:block; position:absolute; z-index:2\"><img src=\"HtmlSaveOptions.RoundtripInformation.001.png\" width=\"227\" height=\"132\" alt=\"\" style=\"margin-top:199.01pt; margin-left:-23pt; -aw-left-pos:-22.5pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:294pt; -aw-wrap-type:none; position:absolute\" />");
        }

        [Test]
        public void RoundtripInformationDefaulValue()
        {
            //Assert that default value is true for HTML and false for MHTML and EPUB.
            var saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            Assert.AreEqual(true, saveOptions.ExportRoundtripInformation);

            saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
            Assert.AreEqual(false, saveOptions.ExportRoundtripInformation);

            saveOptions = new HtmlSaveOptions(SaveFormat.Epub);
            Assert.AreEqual(false, saveOptions.ExportRoundtripInformation);
        }
    }
}