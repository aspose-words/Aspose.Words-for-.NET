// Copyright (c) Aspose 2002-2021. All Rights Reserved.

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
