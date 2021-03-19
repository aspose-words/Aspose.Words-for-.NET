using System.Drawing;
using Aspose.Words;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents
{
    internal class WorkingWithFormFields : DocsExamplesBase
    {
        [Test]
        public void InsertFormFields()
        {
            //ExStart:InsertFormFields
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            string[] items = { "One", "Two", "Three" };
            builder.InsertComboBox("DropDown", items, 0);
            //ExEnd:InsertFormFields
        }

        [Test]
        public void FormFieldsWorkWithProperties()
        {
            //ExStart:FormFieldsWorkWithProperties
            Document doc = new Document(MyDir + "Form fields.docx");
            FormField formField = doc.Range.FormFields[3];

            if (formField.Type == FieldType.FieldFormTextInput)
                formField.Result = "My name is " + formField.Name;
            //ExEnd:FormFieldsWorkWithProperties            
        }

        [Test]
        public void FormFieldsGetFormFieldsCollection()
        {
            //ExStart:FormFieldsGetFormFieldsCollection
            Document doc = new Document(MyDir + "Form fields.docx");
            
            FormFieldCollection formFields = doc.Range.FormFields;
            //ExEnd:FormFieldsGetFormFieldsCollection
        }

        [Test]
        public void FormFieldsGetByName()
        {
            //ExStart:FormFieldsFontFormatting
            //ExStart:FormFieldsGetByName
            Document doc = new Document(MyDir + "Form fields.docx");

            FormFieldCollection documentFormFields = doc.Range.FormFields;

            FormField formField1 = documentFormFields[3];
            FormField formField2 = documentFormFields["Text2"];
            //ExEnd:FormFieldsGetByName

            formField1.Font.Size = 20;
            formField2.Font.Color = Color.Red;
            //ExEnd:FormFieldsFontFormatting
        }
    }
}