// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
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
            Assert.That(ConvertUtil.InchToPoint(1), Is.EqualTo(72.0d));
            Assert.That(ConvertUtil.PointToInch(72), Is.EqualTo(1.0d));

            // Add content to demonstrate the new margins.
            builder.Writeln($"This Text is {pageSetup.LeftMargin} points/{ConvertUtil.PointToInch(pageSetup.LeftMargin)} inches from the left, " +
                            $"{pageSetup.RightMargin} points/{ConvertUtil.PointToInch(pageSetup.RightMargin)} inches from the right, " +
                            $"{pageSetup.TopMargin} points/{ConvertUtil.PointToInch(pageSetup.TopMargin)} inches from the top, " +
                            $"and {pageSetup.BottomMargin} points/{ConvertUtil.PointToInch(pageSetup.BottomMargin)} inches from the bottom of the page.");

            doc.Save(ArtifactsDir + "UtilityClasses.PointsAndInches.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "UtilityClasses.PointsAndInches.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.That(pageSetup.TopMargin, Is.EqualTo(72.0d).Within(0.01d));
            Assert.That(ConvertUtil.PointToInch(pageSetup.TopMargin), Is.EqualTo(1.0d).Within(0.01d));
            Assert.That(pageSetup.BottomMargin, Is.EqualTo(144.0d).Within(0.01d));
            Assert.That(ConvertUtil.PointToInch(pageSetup.BottomMargin), Is.EqualTo(2.0d).Within(0.01d));
            Assert.That(pageSetup.LeftMargin, Is.EqualTo(180.0d).Within(0.01d));
            Assert.That(ConvertUtil.PointToInch(pageSetup.LeftMargin), Is.EqualTo(2.5d).Within(0.01d));
            Assert.That(pageSetup.RightMargin, Is.EqualTo(108.0d).Within(0.01d));
            Assert.That(ConvertUtil.PointToInch(pageSetup.RightMargin), Is.EqualTo(1.5d).Within(0.01d));
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
            Assert.That(ConvertUtil.MillimeterToPoint(10), Is.EqualTo(28.34d).Within(0.01d));

            // Add content to demonstrate the new margins.
            builder.Writeln($"This Text is {pageSetup.LeftMargin} points from the left, " +
                            $"{pageSetup.RightMargin} points from the right, " +
                            $"{pageSetup.TopMargin} points from the top, " +
                            $"and {pageSetup.BottomMargin} points from the bottom of the page.");

            doc.Save(ArtifactsDir + "UtilityClasses.PointsAndMillimeters.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "UtilityClasses.PointsAndMillimeters.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.That(pageSetup.TopMargin, Is.EqualTo(85.05d).Within(0.01d));
            Assert.That(pageSetup.BottomMargin, Is.EqualTo(141.75d).Within(0.01d));
            Assert.That(pageSetup.LeftMargin, Is.EqualTo(226.75d).Within(0.01d));
            Assert.That(pageSetup.RightMargin, Is.EqualTo(113.4d).Within(0.01d));
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
            Assert.That(ConvertUtil.PixelToPoint(1), Is.EqualTo(0.75d));
            Assert.That(ConvertUtil.PointToPixel(0.75), Is.EqualTo(1.0d));

            // The default DPI value used is 96.
            Assert.That(ConvertUtil.PixelToPoint(1, 96), Is.EqualTo(0.75d));

            // Add content to demonstrate the new margins.
            builder.Writeln($"This Text is {pageSetup.LeftMargin} points/{ConvertUtil.PointToPixel(pageSetup.LeftMargin)} pixels from the left, " +
                            $"{pageSetup.RightMargin} points/{ConvertUtil.PointToPixel(pageSetup.RightMargin)} pixels from the right, " +
                            $"{pageSetup.TopMargin} points/{ConvertUtil.PointToPixel(pageSetup.TopMargin)} pixels from the top, " +
                            $"and {pageSetup.BottomMargin} points/{ConvertUtil.PointToPixel(pageSetup.BottomMargin)} pixels from the bottom of the page.");

            doc.Save(ArtifactsDir + "UtilityClasses.PointsAndPixels.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "UtilityClasses.PointsAndPixels.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.That(pageSetup.TopMargin, Is.EqualTo(75.0d).Within(0.01d));
            Assert.That(ConvertUtil.PointToPixel(pageSetup.TopMargin), Is.EqualTo(100.0d).Within(0.01d));
            Assert.That(pageSetup.BottomMargin, Is.EqualTo(150.0d).Within(0.01d));
            Assert.That(ConvertUtil.PointToPixel(pageSetup.BottomMargin), Is.EqualTo(200.0d).Within(0.01d));
            Assert.That(pageSetup.LeftMargin, Is.EqualTo(168.75d).Within(0.01d));
            Assert.That(ConvertUtil.PointToPixel(pageSetup.LeftMargin), Is.EqualTo(225.0d).Within(0.01d));
            Assert.That(pageSetup.RightMargin, Is.EqualTo(93.75d).Within(0.01d));
            Assert.That(ConvertUtil.PointToPixel(pageSetup.RightMargin), Is.EqualTo(125.0d).Within(0.01d));
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

            Assert.That(pageSetup.TopMargin, Is.EqualTo(37.5d).Within(0.01d));

            // At the default DPI of 96, a pixel is 0.75 points.
            Assert.That(ConvertUtil.PixelToPoint(1), Is.EqualTo(0.75d));

            builder.Writeln($"This Text is {pageSetup.TopMargin} points/{ConvertUtil.PointToPixel(pageSetup.TopMargin, myDpi)} " +
                            $"pixels (at a DPI of {myDpi}) from the top of the page.");

            // Set a new DPI and adjust the top margin value accordingly.
            const double newDpi = 300;
            pageSetup.TopMargin = ConvertUtil.PixelToNewDpi(pageSetup.TopMargin, myDpi, newDpi);
            Assert.That(pageSetup.TopMargin, Is.EqualTo(59.0d).Within(0.01d));

            builder.Writeln($"At a DPI of {newDpi}, the text is now {pageSetup.TopMargin} points/{ConvertUtil.PointToPixel(pageSetup.TopMargin, myDpi)} " +
                            "pixels from the top of the page.");

            doc.Save(ArtifactsDir + "UtilityClasses.PointsAndPixelsDpi.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "UtilityClasses.PointsAndPixelsDpi.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.That(pageSetup.TopMargin, Is.EqualTo(59.0d).Within(0.01d));
            Assert.That(ConvertUtil.PointToPixel(pageSetup.TopMargin), Is.EqualTo(78.66).Within(0.01d));
            Assert.That(ConvertUtil.PointToPixel(pageSetup.TopMargin, myDpi), Is.EqualTo(157.33).Within(0.01d));
            Assert.That(ConvertUtil.PointToPixel(100), Is.EqualTo(133.33d).Within(0.01d));
            Assert.That(ConvertUtil.PointToPixel(100, myDpi), Is.EqualTo(266.66d).Within(0.01d));
        }
    }
}