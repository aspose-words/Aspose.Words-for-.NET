// Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Fields;
using NUnit.Framework;
using QA_Tests.Tests;

namespace QA_Tests.Examples.Fields
{
    [TestFixture]
    public class ExFormFields : QaTestsBase
    {
        [Test]
        public void FormFieldsGetFormFieldsCollection()
        {
            //ExStart
            //ExFor:Range.FormFields
            //ExFor:FormFieldCollection
            //ExId:FormFieldsGetFormFieldsCollection
            //ExSummary:Shows how to get a collection of form fields.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "FormFields.doc");
            FormFieldCollection formFields = doc.Range.FormFields;
            //ExEnd
        }

        [Test]
        public void FormFieldsGetByName()
        {
            //ExStart
            //ExFor:FormField
            //ExId:FormFieldsGetByName
            //ExSummary:Shows how to access form fields.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "FormFields.doc");
            FormFieldCollection documentFormFields = doc.Range.FormFields;

            FormField formField1 = documentFormFields[3];
            FormField formField2 = documentFormFields["CustomerName"];
            //ExEnd
        }

        [Test]
        public void FormFieldsWorkWithProperties()
        {
            //ExStart
            //ExFor:FormField
            //ExFor:FormField.Result
            //ExFor:FormField.Type
            //ExFor:FormField.Name
            //ExId:FormFieldsWorkWithProperties
            //ExSummary:Shows how to work with form field name, type, and result.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "FormFields.doc");
            
            FormField formField = doc.Range.FormFields[3];

            if (formField.Type.Equals(FieldType.FieldFormTextInput))
                formField.Result = "My name is " + formField.Name;
            //ExEnd
        }

        [Test]
        public void InsertAndRetrieveFormFields()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertTextInput
            //ExId:FormFieldsInsertAndRetrieve
            //ExSummary:Shows how to insert FormFields, set options annd gather them back in for use 
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a text input field. The unique name of this field is "TextInput1", the other parameters define
            // what type of FormField it is, the format of the text, the field result and the maximum text length (0 = no limit)
            builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "", "", 0);
            //ExEnd
        }
    }
}
