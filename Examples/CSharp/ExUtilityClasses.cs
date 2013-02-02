//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using Aspose.Words;
using NUnit.Framework;

namespace Examples
{
    [TestFixture]
    public class ExUtilityClasses : ExBase
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
    }
}
