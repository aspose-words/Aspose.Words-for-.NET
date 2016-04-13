// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;

using Aspose.Words;
using Aspose.Words.Fields;

using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExFormFields : ApiExampleBase
    {
        [Test]
        public void FormFieldsGetFormFieldsCollection()
        {
            //ExStart
            //ExFor:Range.FormFields
            //ExFor:FormFieldCollection
            //ExId:FormFieldsGetFormFieldsCollection
            //ExSummary:Shows how to get a collection of form fields.
            Document doc = new Document(MyDir + "FormFields.doc");
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
            Document doc = new Document(MyDir + "FormFields.doc");
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
            Document doc = new Document(MyDir + "FormFields.doc");
            
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
            //ExSummary:Shows how to insert form fields, set options and gather them back in for use 
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a text input field. The unique name of this field is "TextInput1", the other parameters define
            // what type of FormField it is, the format of the text, the field result and the maximum text length (0 = no limit)
            builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "", "", 0);
            //ExEnd
        }

        [Test]
        public void DeleteFormField()
        {
            //ExStart
            //ExFor:FormField.RemoveField
            //ExSummary:Shows how to delete complete form field
            Document doc = new Document(MyDir + "FormFields.doc");

            FormField formField = doc.Range.FormFields[3];
            formField.RemoveField();
            //ExEnd

            FormField formFieldAfter = doc.Range.FormFields[3];
            
            Assert.IsNull(formFieldAfter);
        }

        [Test]
        public void DeleteFormFieldAssociatedWithTheFormField()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartBookmark("MyBookmark");
            builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "TestFormField", "SomeText", 0);
            builder.EndBookmark("MyBookmark");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            BookmarkCollection bookmarkBeforeDeleteFormField = doc.Range.Bookmarks;
            Assert.AreEqual("MyBookmark", bookmarkBeforeDeleteFormField[0].Name);

            FormField formField = doc.Range.FormFields[0];
            formField.RemoveField();

            BookmarkCollection bookmarkAfterDeleteFormField = doc.Range.Bookmarks;
            Assert.AreEqual("MyBookmark", bookmarkAfterDeleteFormField[0].Name);
        }
    }
}
