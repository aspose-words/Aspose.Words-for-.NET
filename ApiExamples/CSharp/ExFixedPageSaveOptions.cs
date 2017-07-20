// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;
using System.Collections.Generic;

namespace ApiExamples
{
    [TestFixture]
    internal class ExFixedPageSaveOptions : ApiExampleBase
    {
        public static IEnumerable<TestCaseData> FixedPageSaveOptionsDefaultValuesData
        {
            get
            {
                yield return new TestCaseData(new HtmlFixedSaveOptions());
                yield return new TestCaseData(new ImageSaveOptions(SaveFormat.Jpeg));
                yield return new TestCaseData(new PdfSaveOptions());
                yield return new TestCaseData(new PsSaveOptions());
                yield return new TestCaseData(new SvgSaveOptions());
                yield return new TestCaseData(new XamlFixedSaveOptions());
                yield return new TestCaseData(new XpsSaveOptions());
                yield return new TestCaseData(new SwfSaveOptions());
            }
        }

        public static IEnumerable<TestCaseData> FixedPageSaveOptionsData
        {
            get
            {
                yield return new TestCaseData(new HtmlFixedSaveOptions(), 100, NumeralFormat.ArabicIndic, int.MaxValue, 1, EmfPlusDualRenderingMode.Emf, false, MetafileRenderingMode.Vector, false, true);
                yield return new TestCaseData(new ImageSaveOptions(SaveFormat.Jpeg), 100, NumeralFormat.ArabicIndic, int.MaxValue, 1, EmfPlusDualRenderingMode.Emf, false, MetafileRenderingMode.Vector, false, true);
                yield return new TestCaseData(new PdfSaveOptions(), 100, NumeralFormat.ArabicIndic, int.MaxValue, 1, EmfPlusDualRenderingMode.Emf, false, MetafileRenderingMode.Vector, false, true);
                yield return new TestCaseData(new PsSaveOptions(), 100, NumeralFormat.ArabicIndic, int.MaxValue, 1, EmfPlusDualRenderingMode.Emf, false, MetafileRenderingMode.Vector, false, true);
                yield return new TestCaseData(new SvgSaveOptions(), 100, NumeralFormat.ArabicIndic, int.MaxValue, 1, EmfPlusDualRenderingMode.Emf, false, MetafileRenderingMode.Vector, false, true);
                yield return new TestCaseData(new XamlFixedSaveOptions(), 100, NumeralFormat.ArabicIndic, int.MaxValue, 1, EmfPlusDualRenderingMode.Emf, false, MetafileRenderingMode.Vector, false, true);
                yield return new TestCaseData(new XpsSaveOptions(), 100, NumeralFormat.ArabicIndic, int.MaxValue, 1, EmfPlusDualRenderingMode.Emf, false, MetafileRenderingMode.Vector, false, true);
                yield return new TestCaseData(new SwfSaveOptions(), 100, NumeralFormat.ArabicIndic, int.MaxValue, 1, EmfPlusDualRenderingMode.Emf, false, MetafileRenderingMode.Vector, false, true);
            }
        }

        [Test]
        [Ignore("Bug?")]
        [TestCaseSource("FixedPageSaveOptionsDefaultValuesData")]
        public void FixedPageSaveOptionsDefaultValues(FixedPageSaveOptions objectSaveOptions)
        {
            FixedPageSaveOptions saveOptions = objectSaveOptions;

            Assert.AreEqual(objectSaveOptions.GetType().Name == "PdfSaveOptions" ? 100 : 95, saveOptions.JpegQuality);
            Assert.AreEqual(NumeralFormat.European, saveOptions.NumeralFormat);
            Assert.AreEqual(int.MaxValue, saveOptions.PageCount);
            Assert.AreEqual(0, saveOptions.PageIndex);
            Assert.AreEqual(EmfPlusDualRenderingMode.EmfPlusWithFallback, saveOptions.MetafileRenderingOptions.EmfPlusDualRenderingMode);
            Assert.AreEqual(true, saveOptions.MetafileRenderingOptions.EmulateRasterOperations);
            Assert.AreEqual(objectSaveOptions.GetType().Name == "ImageSaveOptions" ? MetafileRenderingMode.Bitmap : MetafileRenderingMode.VectorWithFallback, saveOptions.MetafileRenderingOptions.RenderingMode);
            Assert.AreEqual(true, saveOptions.MetafileRenderingOptions.UseEmfEmbeddedToWmf);
            Assert.AreEqual(false, saveOptions.OptimizeOutput); //bug?
        }

        [Test]
        [TestCaseSource("FixedPageSaveOptionsData")]
        public void SaveInFixedFormat(FixedPageSaveOptions objectSaveOptions, int jpegQuality, NumeralFormat numeralFormat, int pageCount, int pageIndex, EmfPlusDualRenderingMode emfPlusDualRenderingMode, bool emulateRasterOperations, MetafileRenderingMode metafileRendering, bool useEmfEmbeddedToWmf, bool optimizeOutput)
        {
            FixedPageSaveOptions saveOptions = objectSaveOptions;

            saveOptions.JpegQuality = jpegQuality;
            saveOptions.NumeralFormat = numeralFormat;
            saveOptions.PageCount = pageCount;
            saveOptions.PageIndex = pageIndex;
            saveOptions.MetafileRenderingOptions.EmfPlusDualRenderingMode = emfPlusDualRenderingMode;
            saveOptions.MetafileRenderingOptions.EmulateRasterOperations = emulateRasterOperations;
            saveOptions.MetafileRenderingOptions.RenderingMode = metafileRendering;
            saveOptions.MetafileRenderingOptions.UseEmfEmbeddedToWmf = useEmfEmbeddedToWmf;
            saveOptions.OptimizeOutput = optimizeOutput;

            Assert.AreEqual(jpegQuality, saveOptions.JpegQuality);
            Assert.AreEqual(numeralFormat, saveOptions.NumeralFormat);
            Assert.AreEqual(pageCount, saveOptions.PageCount);
            Assert.AreEqual(pageIndex, saveOptions.PageIndex);
            Assert.AreEqual(emfPlusDualRenderingMode, saveOptions.MetafileRenderingOptions.EmfPlusDualRenderingMode);
            Assert.AreEqual(emulateRasterOperations, saveOptions.MetafileRenderingOptions.EmulateRasterOperations);
            Assert.AreEqual(metafileRendering, saveOptions.MetafileRenderingOptions.RenderingMode);
            Assert.AreEqual(useEmfEmbeddedToWmf, saveOptions.MetafileRenderingOptions.UseEmfEmbeddedToWmf);
            Assert.AreEqual(optimizeOutput, saveOptions.OptimizeOutput);
        }
    }
}