// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExUtilityClasses : ApiExampleBase
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
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            PageSetup pageSetup = builder.PageSetup;
            pageSetup.TopMargin = ConvertUtil.InchToPoint(1.0);
            pageSetup.BottomMargin = ConvertUtil.InchToPoint(1.0);
            pageSetup.LeftMargin = ConvertUtil.InchToPoint(1.5);
            pageSetup.RightMargin = ConvertUtil.InchToPoint(1.5);
            pageSetup.HeaderDistance = ConvertUtil.InchToPoint(0.2);
            pageSetup.FooterDistance = ConvertUtil.InchToPoint(0.2);
            //ExEnd
        }

        [Test]
        public void MillimeterToPointEx()
        {
            //ExStart
            //ExFor:ConvertUtil.MillimeterToPoint
            //ExSummary:Shows how to specify page properties in millimeters.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            PageSetup pageSetup = builder.PageSetup;
            pageSetup.TopMargin = ConvertUtil.MillimeterToPoint(25.0);
            pageSetup.BottomMargin = ConvertUtil.MillimeterToPoint(25.0);
            pageSetup.LeftMargin = ConvertUtil.MillimeterToPoint(37.5);
            pageSetup.RightMargin = ConvertUtil.MillimeterToPoint(37.5);
            pageSetup.HeaderDistance = ConvertUtil.MillimeterToPoint(5.0);
            pageSetup.FooterDistance = ConvertUtil.MillimeterToPoint(5.0);

            builder.Writeln("Hello world.");
            builder.Document.Save(MyDir + @"\Artifacts\PageSetup.PageMargins.doc");
            //ExEnd
        }

        [Test]
        public void PointToInchEx()
        {
            //ExStart
            //ExFor:ConvertUtil.PointToInch
            //ExSummary:Shows how to convert points to inches.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            PageSetup pageSetup = builder.PageSetup;
            pageSetup.TopMargin = ConvertUtil.InchToPoint(2.0);

            Console.WriteLine("The size of my top margin is {0} points, or {1} inches.",
                pageSetup.TopMargin, ConvertUtil.PointToInch(pageSetup.TopMargin));
            //ExEnd
        }

        [Test]
        public void PixelToPointEx()
        {
            //ExStart
            //ExFor:ConvertUtil.PixelToPoint(double)
            //ExFor:ConvertUtil.PixelToPoint(double, double)
            //ExSummary:Shows how to specify page properties in pixels with default and custom resolution.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            PageSetup pageSetupNoDpi = builder.PageSetup;
            pageSetupNoDpi.TopMargin = ConvertUtil.PixelToPoint(100.0);
            pageSetupNoDpi.BottomMargin = ConvertUtil.PixelToPoint(100.0);
            pageSetupNoDpi.LeftMargin = ConvertUtil.PixelToPoint(150.0);
            pageSetupNoDpi.RightMargin = ConvertUtil.PixelToPoint(150.0);
            pageSetupNoDpi.HeaderDistance = ConvertUtil.PixelToPoint(20.0);
            pageSetupNoDpi.FooterDistance = ConvertUtil.PixelToPoint(20.0);

            builder.Writeln("Hello world.");
            builder.Document.Save(MyDir + @"\Artifacts\PageSetup.PageMargins.DefaultResolution.doc");

            double myDpi = 150.0;

            PageSetup pageSetupWithDpi = builder.PageSetup;
            pageSetupWithDpi.TopMargin = ConvertUtil.PixelToPoint(100.0, myDpi);
            pageSetupWithDpi.BottomMargin = ConvertUtil.PixelToPoint(100.0, myDpi);
            pageSetupWithDpi.LeftMargin = ConvertUtil.PixelToPoint(150.0, myDpi);
            pageSetupWithDpi.RightMargin = ConvertUtil.PixelToPoint(150.0, myDpi);
            pageSetupWithDpi.HeaderDistance = ConvertUtil.PixelToPoint(20.0, myDpi);
            pageSetupWithDpi.FooterDistance = ConvertUtil.PixelToPoint(20.0, myDpi);

            builder.Document.Save(MyDir + @"\Artifacts\PageSetup.PageMargins.CustomResolution.doc");
            //ExEnd
        }

        [Test]
        public void PointToPixelEx()
        {
            //ExStart
            //ExFor:ConvertUtil.PointToPixel(double)
            //ExFor:ConvertUtil.PointToPixel(double, double)
            //ExSummary:Shows how to use convert points to pixels with default and custom resolution.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            PageSetup pageSetup = builder.PageSetup;
            pageSetup.TopMargin = ConvertUtil.PixelToPoint(2.0);

            double myDpi = 192.0;

            Console.WriteLine("The size of my top margin is {0} points, or {1} pixels with default resolution.",
                pageSetup.TopMargin, ConvertUtil.PointToPixel(pageSetup.TopMargin));

            Console.WriteLine("The size of my top margin is {0} points, or {1} pixels with custom resolution.",
                pageSetup.TopMargin, ConvertUtil.PointToPixel(pageSetup.TopMargin, myDpi));
            //ExEnd
        }

        [Test]
        public void PixelToNewDpiEx()
        {
            //ExStart
            //ExFor:ConvertUtil.PixelToNewDpi
            //ExSummary:Shows how to check how an amount of pixels changes when the dpi is changed.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            PageSetup pageSetup = builder.PageSetup;
            pageSetup.TopMargin = 72;
            double oldDpi = 92.0;
            double newDpi = 192.0;

            Console.WriteLine("{0} pixels at {1} dpi becomes {2} pixels at {3} dpi.",
                pageSetup.TopMargin, oldDpi, ConvertUtil.PixelToNewDpi(pageSetup.TopMargin, oldDpi, newDpi), newDpi);
            //ExEnd
        }
    }
}
