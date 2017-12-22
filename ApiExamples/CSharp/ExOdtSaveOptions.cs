// Copyright (c) 2001-2017 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    internal class ExOdtSaveOptions : ApiExampleBase
    {
        [Test]
        public void MeasureUnitOption()
        {
            //ExStart
            //ExFor:OdtSaveOptions.MeasureUnit
            //ExSummary: Show how to work with units of measure of document content
            Document doc = new Document(MyDir + "OdtSaveOptions.MeasureUnit.docx");

            //Open Office uses centimeters, MS Office uses inches
            OdtSaveOptions saveOptions = new OdtSaveOptions();
            saveOptions.MeasureUnit = OdtSaveMeasureUnit.Inches;

            doc.Save(MyDir + @"\Artifacts\OdtSaveOptions.MeasureUnit.odt");
            //ExEnd
        }
    }
}