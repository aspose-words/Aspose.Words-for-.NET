// Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using NUnit.Framework;
using QA_Tests.Tests;

namespace QA_Tests.Examples.ConvertUtil
{
    [TestFixture]
    public class ExUtilityClasses : QaTestsBase
    {
        [Test]
        public void UtilityClassesUseControlCharacters()
        {
            string text = "test\r";
            //ExStart
            //ExFor:ControlChar
            //ExFor:ControlChar.Cr
            //ExFor:ControlChar.CrLf
            //ExId:UtilityClassesUseControlCharacters
            //ExSummary:Shows how to use control characters.
            // Replace "\r" control character with "\r\n"
            text = text.Replace(ControlChar.Cr, ControlChar.CrLf);
            //ExEnd
        }

        [Test]
        public void UtilityClassesConvertBetweenMeasurementUnits()
        {
            //ExStart
            //ExFor:ConvertUtil
            //ExId:UtilityClassesConvertBetweenMeasurementUnits
            //ExSummary:Shows how to specify page properties in inches.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Aspose.Words.PageSetup pageSetup = builder.PageSetup;
            pageSetup.TopMargin = Aspose.Words.ConvertUtil.InchToPoint(1.0);
            pageSetup.BottomMargin = Aspose.Words.ConvertUtil.InchToPoint(1.0);
            pageSetup.LeftMargin = Aspose.Words.ConvertUtil.InchToPoint(1.5);
            pageSetup.RightMargin = Aspose.Words.ConvertUtil.InchToPoint(1.5);
            pageSetup.HeaderDistance = Aspose.Words.ConvertUtil.InchToPoint(0.2);
            pageSetup.FooterDistance = Aspose.Words.ConvertUtil.InchToPoint(0.2);
            //ExEnd
        }

        [Test]
        public void MillimeterToPointEx()
        {
            //ExStart
            //ExFor:MillimeterToPoint
            //ExId:MillimeterToPointEx
            //ExSummary:Shows how to specify page properties in millimeters.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Aspose.Words.PageSetup pageSetup = builder.PageSetup;
            pageSetup.TopMargin = Aspose.Words.ConvertUtil.MillimeterToPoint(25.0);
            pageSetup.BottomMargin = Aspose.Words.ConvertUtil.MillimeterToPoint(25.0);
            pageSetup.LeftMargin = Aspose.Words.ConvertUtil.MillimeterToPoint(37.5);
            pageSetup.RightMargin = Aspose.Words.ConvertUtil.MillimeterToPoint(37.5);
            pageSetup.HeaderDistance = Aspose.Words.ConvertUtil.MillimeterToPoint(5.0);
            pageSetup.FooterDistance = Aspose.Words.ConvertUtil.MillimeterToPoint(5.0);

            builder.Writeln("Hello world.");
            builder.Document.Save(ExDir + "PageSetup.PageMargins Out.doc");
            //ExEnd
        }

        [Test]
        public void PointToInchEx()
        {
            //ExStart
            //ExFor:ConvertUtil.PointToInch
            //ExSummary:Shows how to use PointToInch.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Aspose.Words.PageSetup pageSetup = builder.PageSetup;
            pageSetup.TopMargin = Aspose.Words.ConvertUtil.InchToPoint(2.0);

            Console.WriteLine("The size of my top margin is {0} points, or {1} inches.",
                pageSetup.TopMargin, Aspose.Words.ConvertUtil.PointToInch(pageSetup.TopMargin));
            //ExEnd
        }
    }
}
