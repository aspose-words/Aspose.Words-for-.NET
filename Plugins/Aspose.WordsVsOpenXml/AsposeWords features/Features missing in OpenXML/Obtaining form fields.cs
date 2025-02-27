// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features.Features_missing_in_OpenXML
{
    [TestFixture]
    public class ObtainingFormFields : TestUtil
    {
        [Test]
        public static void ObtainingFormFieldsFeature()
        {
            //Shows how to get a collection of form fields.
            Document doc = new Document(MyDir + "Obtaining form fields.docx");
            
            FormFieldCollection formFields = doc.Range.FormFields;
            FormField formField1 = formFields[3];
            FormField formField2 = formFields["CustomerName"];
        }
    }
}
