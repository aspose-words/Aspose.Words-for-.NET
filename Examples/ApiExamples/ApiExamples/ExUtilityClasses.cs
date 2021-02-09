// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExUtilityClasses : ApiExampleBase
    {
        [Test]
        public void PointsAndInches()
        {
            //ExStart
            //ExFor:ConvertUtil
            //ExFor:ConvertUtil.PointToInch
            //ExFor:ConvertUtil.InchToPoint
            //ExSummary:Shows how to specify page properties in inches.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // A section's "Page Setup" defines the size of the page margins in points.
            // We can also use the "ConvertUtil" class to use a more familiar measurement unit,
            // such as inches when defining boundaries.
            PageSetup pageSetup = builder.PageSetup;
            pageSetup.TopMargin = ConvertUtil.InchToPoint(1.0);
            pageSetup.BottomMargin = ConvertUtil.InchToPoint(2.0);
            pageSetup.LeftMargin = ConvertUtil.InchToPoint(2.5);
            pageSetup.RightMargin = ConvertUtil.InchToPoint(1.5);

            // An inch is 72 points.
            Assert.AreEqual(72.0d, ConvertUtil.InchToPoint(1));
            Assert.AreEqual(1.0d, ConvertUtil.PointToInch(72));

            // Add content to demonstrate the new margins.
            builder.Writeln($"This Text is {pageSetup.LeftMargin} points/{ConvertUtil.PointToInch(pageSetup.LeftMargin)} inches from the left, " +
                            $"{pageSetup.RightMargin} points/{ConvertUtil.PointToInch(pageSetup.RightMargin)} inches from the right, " +
                            $"{pageSetup.TopMargin} points/{ConvertUtil.PointToInch(pageSetup.TopMargin)} inches from the top, " +
                            $"and {pageSetup.BottomMargin} points/{ConvertUtil.PointToInch(pageSetup.BottomMargin)} inches from the bottom of the page.");

            doc.Save(ArtifactsDir + "UtilityClasses.PointsAndInches.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "UtilityClasses.PointsAndInches.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.AreEqual(72.0d, pageSetup.TopMargin, 0.01d);
            Assert.AreEqual(1.0d, ConvertUtil.PointToInch(pageSetup.TopMargin), 0.01d);
            Assert.AreEqual(144.0d, pageSetup.BottomMargin, 0.01d);
            Assert.AreEqual(2.0d, ConvertUtil.PointToInch(pageSetup.BottomMargin), 0.01d);
            Assert.AreEqual(180.0d, pageSetup.LeftMargin, 0.01d);
            Assert.AreEqual(2.5d, ConvertUtil.PointToInch(pageSetup.LeftMargin), 0.01d);
            Assert.AreEqual(108.0d, pageSetup.RightMargin, 0.01d);
            Assert.AreEqual(1.5d, ConvertUtil.PointToInch(pageSetup.RightMargin), 0.01d);
        }

        [Test]
        public void PointsAndMillimeters()
        {
            //ExStart
            //ExFor:ConvertUtil.MillimeterToPoint
            //ExSummary:Shows how to specify page properties in millimeters.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // A section's "Page Setup" defines the size of the page margins in points.
            // We can also use the "ConvertUtil" class to use a more familiar measurement unit,
            // such as millimeters when defining boundaries.
            PageSetup pageSetup = builder.PageSetup;
            pageSetup.TopMargin = ConvertUtil.MillimeterToPoint(30);
            pageSetup.BottomMargin = ConvertUtil.MillimeterToPoint(50);
            pageSetup.LeftMargin = ConvertUtil.MillimeterToPoint(80);
            pageSetup.RightMargin = ConvertUtil.MillimeterToPoint(40);

            // A centimeter is approximately 28.3 points.
            Assert.AreEqual(28.34d, ConvertUtil.MillimeterToPoint(10), 0.01d);

            // Add content to demonstrate the new margins.
            builder.Writeln($"This Text is {pageSetup.LeftMargin} points from the left, " +
                            $"{pageSetup.RightMargin} points from the right, " +
                            $"{pageSetup.TopMargin} points from the top, " +
                            $"and {pageSetup.BottomMargin} points from the bottom of the page.");

            doc.Save(ArtifactsDir + "UtilityClasses.PointsAndMillimeters.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "UtilityClasses.PointsAndMillimeters.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.AreEqual(85.05d, pageSetup.TopMargin, 0.01d);
            Assert.AreEqual(141.75d, pageSetup.BottomMargin, 0.01d);
            Assert.AreEqual(226.75d, pageSetup.LeftMargin, 0.01d);
            Assert.AreEqual(113.4d, pageSetup.RightMargin, 0.01d);
        }

        [Test]
        public void PointsAndPixels()
        {
            //ExStart
            //ExFor:ConvertUtil.PixelToPoint(double)
            //ExFor:ConvertUtil.PointToPixel(double)
            //ExSummary:Shows how to specify page properties in pixels.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // A section's "Page Setup" defines the size of the page margins in points.
            // We can also use the "ConvertUtil" class to use a different measurement unit,
            // such as pixels when defining boundaries.
            PageSetup pageSetup = builder.PageSetup;
            pageSetup.TopMargin = ConvertUtil.PixelToPoint(100);
            pageSetup.BottomMargin = ConvertUtil.PixelToPoint(200);
            pageSetup.LeftMargin = ConvertUtil.PixelToPoint(225);
            pageSetup.RightMargin = ConvertUtil.PixelToPoint(125);

            // A pixel is 0.75 points.
            Assert.AreEqual(0.75d, ConvertUtil.PixelToPoint(1));
            Assert.AreEqual(1.0d, ConvertUtil.PointToPixel(0.75));

            // The default DPI value used is 96.
            Assert.AreEqual(0.75d, ConvertUtil.PixelToPoint(1, 96));

            // Add content to demonstrate the new margins.
            builder.Writeln($"This Text is {pageSetup.LeftMargin} points/{ConvertUtil.PointToPixel(pageSetup.LeftMargin)} pixels from the left, " +
                            $"{pageSetup.RightMargin} points/{ConvertUtil.PointToPixel(pageSetup.RightMargin)} pixels from the right, " +
                            $"{pageSetup.TopMargin} points/{ConvertUtil.PointToPixel(pageSetup.TopMargin)} pixels from the top, " +
                            $"and {pageSetup.BottomMargin} points/{ConvertUtil.PointToPixel(pageSetup.BottomMargin)} pixels from the bottom of the page.");

            doc.Save(ArtifactsDir + "UtilityClasses.PointsAndPixels.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "UtilityClasses.PointsAndPixels.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.AreEqual(75.0d, pageSetup.TopMargin, 0.01d);
            Assert.AreEqual(100.0d, ConvertUtil.PointToPixel(pageSetup.TopMargin), 0.01d);
            Assert.AreEqual(150.0d, pageSetup.BottomMargin, 0.01d);
            Assert.AreEqual(200.0d, ConvertUtil.PointToPixel(pageSetup.BottomMargin), 0.01d);
            Assert.AreEqual(168.75d, pageSetup.LeftMargin, 0.01d);
            Assert.AreEqual(225.0d, ConvertUtil.PointToPixel(pageSetup.LeftMargin), 0.01d);
            Assert.AreEqual(93.75d, pageSetup.RightMargin, 0.01d);
            Assert.AreEqual(125.0d, ConvertUtil.PointToPixel(pageSetup.RightMargin), 0.01d);
        }

        [Test]
        public void PointsAndPixelsDpi()
        {
            //ExStart
            //ExFor:ConvertUtil.PixelToNewDpi
            //ExFor:ConvertUtil.PixelToPoint(double, double)
            //ExFor:ConvertUtil.PointToPixel(double, double)
            //ExSummary:Shows how to use convert points to pixels with default and custom resolution.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Define the size of the top margin of this section in pixels, according to a custom DPI.
            const double myDpi = 192;

            PageSetup pageSetup = builder.PageSetup;
            pageSetup.TopMargin = ConvertUtil.PixelToPoint(100, myDpi);

            Assert.AreEqual(37.5d, pageSetup.TopMargin, 0.01d);

            // At the default DPI of 96, a pixel is 0.75 points.
            Assert.AreEqual(0.75d, ConvertUtil.PixelToPoint(1));

            builder.Writeln($"This Text is {pageSetup.TopMargin} points/{ConvertUtil.PointToPixel(pageSetup.TopMargin, myDpi)} " +
                            $"pixels (at a DPI of {myDpi}) from the top of the page.");

            // Set a new DPI and adjust the top margin value accordingly.
            const double newDpi = 300;
            pageSetup.TopMargin = ConvertUtil.PixelToNewDpi(pageSetup.TopMargin, myDpi, newDpi);
            Assert.AreEqual(59.0d, pageSetup.TopMargin, 0.01d);

            builder.Writeln($"At a DPI of {newDpi}, the text is now {pageSetup.TopMargin} points/{ConvertUtil.PointToPixel(pageSetup.TopMargin, myDpi)} " +
                            "pixels from the top of the page.");

            doc.Save(ArtifactsDir + "UtilityClasses.PointsAndPixelsDpi.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "UtilityClasses.PointsAndPixelsDpi.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.AreEqual(59.0d, pageSetup.TopMargin, 0.01d);
            Assert.AreEqual(78.66, ConvertUtil.PointToPixel(pageSetup.TopMargin), 0.01d);
            Assert.AreEqual(157.33, ConvertUtil.PointToPixel(pageSetup.TopMargin, myDpi), 0.01d);
            Assert.AreEqual(133.33d, ConvertUtil.PointToPixel(100), 0.01d);
            Assert.AreEqual(266.66d, ConvertUtil.PointToPixel(100, myDpi), 0.01d);
        }
    }
}