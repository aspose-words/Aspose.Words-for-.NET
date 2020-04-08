// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExFormFields : ApiExampleBase
    {
        [Test]
        public void FormFieldsWorkWithProperties()
        {
            //ExStart
            //ExFor:FormField
            //ExFor:FormField.Result
            //ExFor:FormField.Type
            //ExFor:FormField.Name
            //ExSummary:Shows how to work with form field name, type, and result.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a DocumentBuilder to insert a combo box form field
            FormField comboBox = builder.InsertComboBox("MyComboBox", new[] { "One", "Two", "Three" }, 0);

            // Verify some of our form field's attributes
            Assert.AreEqual("MyComboBox", comboBox.Name);
            Assert.AreEqual(FieldType.FieldFormDropDown, comboBox.Type);
            Assert.AreEqual("One", comboBox.Result);
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            comboBox = doc.Range.FormFields[0];

            Assert.AreEqual("MyComboBox", comboBox.Name);
            Assert.AreEqual(FieldType.FieldFormDropDown, comboBox.Type);
            Assert.AreEqual("One", comboBox.Result);
        }

        [Test]
        public void InsertAndRetrieveFormFields()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertTextInput
            //ExSummary:Shows how to insert form fields, set options and gather them back in for use.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a text input field. The unique name of this field is "TextInput1", the other parameters define
            // what type of FormField it is, the format of the text, the field result and the maximum text length (0 = no limit)
            builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "", "", 0);
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            FormField textInput = doc.Range.FormFields[0];

            Assert.AreEqual("TextInput1", textInput.Name);
            Assert.AreEqual(TextFormFieldType.Regular, textInput.TextInputType);
            Assert.AreEqual(String.Empty, textInput.TextInputFormat);
            Assert.AreEqual(String.Empty, textInput.Result);
            Assert.AreEqual(0, textInput.MaxLength);
        }

        [Test]
        public void DeleteFormField()
        {
            //ExStart
            //ExFor:FormField.RemoveField
            //ExSummary:Shows how to delete complete form field.
            Document doc = new Document(MyDir + "Form fields.docx");

            FormField formField = doc.Range.FormFields[3];
            formField.RemoveField();
            //ExEnd

            FormField formFieldAfter = doc.Range.FormFields[3];

            Assert.IsNull(formFieldAfter);
        }

        [Test]
        public void DeleteFormFieldAssociatedWithBookmark()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartBookmark("MyBookmark");
            builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "TestFormField", "SomeText", 0);
            builder.EndBookmark("MyBookmark");

            doc = DocumentHelper.SaveOpen(doc);

            BookmarkCollection bookmarkBeforeDeleteFormField = doc.Range.Bookmarks;
            Assert.AreEqual("MyBookmark", bookmarkBeforeDeleteFormField[0].Name);

            FormField formField = doc.Range.FormFields[0];
            formField.RemoveField();

            BookmarkCollection bookmarkAfterDeleteFormField = doc.Range.Bookmarks;
            Assert.AreEqual("MyBookmark", bookmarkAfterDeleteFormField[0].Name);
        }

        //ExStart
        //ExFor:FormField.Accept(DocumentVisitor)
        //ExFor:FormField.CalculateOnExit
        //ExFor:FormField.CheckBoxSize
        //ExFor:FormField.Checked
        //ExFor:FormField.Default
        //ExFor:FormField.DropDownItems
        //ExFor:FormField.DropDownSelectedIndex
        //ExFor:FormField.Enabled
        //ExFor:FormField.EntryMacro
        //ExFor:FormField.ExitMacro
        //ExFor:FormField.HelpText
        //ExFor:FormField.IsCheckBoxExactSize
        //ExFor:FormField.MaxLength
        //ExFor:FormField.OwnHelp
        //ExFor:FormField.OwnStatus
        //ExFor:FormField.SetTextInputValue(Object)
        //ExFor:FormField.StatusText
        //ExFor:FormField.TextInputDefault
        //ExFor:FormField.TextInputFormat
        //ExFor:FormField.TextInputType
        //ExFor:FormFieldCollection
        //ExFor:FormFieldCollection.Clear
        //ExFor:FormFieldCollection.Count
        //ExFor:FormFieldCollection.GetEnumerator
        //ExFor:FormFieldCollection.Item(Int32)
        //ExFor:FormFieldCollection.Item(String)
        //ExFor:FormFieldCollection.Remove(String)
        //ExFor:FormFieldCollection.RemoveAt(Int32)
        //ExFor:Range.FormFields
        //ExSummary:Shows how insert different kinds of form fields into a document and process them with a visitor implementation.
        [Test] //ExSkip
        public void FormField()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a document builder to insert a combo box
            FormField comboBox = builder.InsertComboBox("MyComboBox", new[] { "One", "Two", "Three" }, 0);
            comboBox.CalculateOnExit = true;
            Assert.AreEqual(3, comboBox.DropDownItems.Count);
            Assert.AreEqual(0, comboBox.DropDownSelectedIndex);
            Assert.True(comboBox.Enabled);

            // Use a document builder to insert a check box
            FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 50);
            checkBox.IsCheckBoxExactSize = true;
            checkBox.HelpText = "Right click to check this box";
            checkBox.OwnHelp = true;
            checkBox.StatusText = "Checkbox status text";
            checkBox.OwnStatus = true;
            Assert.AreEqual(50.0d, checkBox.CheckBoxSize);
            Assert.False(checkBox.Checked);
            Assert.False(checkBox.Default);

            builder.Writeln();

            // Use a document builder to insert text input form field
            FormField textInput = builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Your text goes here", 50);
            textInput.EntryMacro = "EntryMacro";
            textInput.ExitMacro = "ExitMacro";
            textInput.TextInputDefault = "Regular";
            textInput.TextInputFormat = "FIRST CAPITAL";
            textInput.SetTextInputValue("This value overrides the one we set during initialization");
            Assert.AreEqual(TextFormFieldType.Regular, textInput.TextInputType);
            Assert.AreEqual(50, textInput.MaxLength);

            // Get the collection of form fields that has accumulated in our document
            FormFieldCollection formFields = doc.Range.FormFields;
            Assert.AreEqual(3, formFields.Count);

            // Our form fields are represented as fields, with field codes FORMDROPDOWN, FORMCHECKBOX and FORMTEXT respectively,
            // made visible by pressing Alt + F9 in Microsoft Word
            // These fields have no switches and the content of their form fields is fully governed by members of the FormField object
            Assert.AreEqual(3, doc.Range.Fields.Count);

            // Iterate over the collection with an enumerator, accepting a visitor with each form field
            FormFieldVisitor formFieldVisitor = new FormFieldVisitor();

            using (IEnumerator<FormField> fieldEnumerator = formFields.GetEnumerator())
                while (fieldEnumerator.MoveNext())
                    fieldEnumerator.Current.Accept(formFieldVisitor);

            Console.WriteLine(formFieldVisitor.GetText());

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.FormField.docx");
            TestFormField(doc); //ExSkip
        }

        /// <summary>
        /// Visitor implementation that prints information about visited form fields. 
        /// </summary>
        public class FormFieldVisitor : DocumentVisitor
        {
            public FormFieldVisitor()
            {
                mBuilder = new StringBuilder();
            }

            /// <summary>
            /// Called when a FormField node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFormField(FormField formField)
            {
                AppendLine(formField.Type + ": \"" + formField.Name + "\"");
                AppendLine("\tStatus: " + (formField.Enabled ? "Enabled" : "Disabled"));
                AppendLine("\tHelp Text:  " + formField.HelpText);
                AppendLine("\tEntry macro name: " + formField.EntryMacro);
                AppendLine("\tExit macro name: " + formField.ExitMacro);

                switch (formField.Type)
                {
                    case FieldType.FieldFormDropDown:
                        AppendLine("\tDrop down items count: " + formField.DropDownItems.Count + ", default selected item index: " + formField.DropDownSelectedIndex);
                        AppendLine("\tDrop down items: " + string.Join(", ", formField.DropDownItems.ToArray()));
                        break;
                    case FieldType.FieldFormCheckBox:
                        AppendLine("\tCheckbox size: " + formField.CheckBoxSize);
                        AppendLine("\t" + "Checkbox is currently: " + (formField.Checked ? "checked, " : "unchecked, ") + "by default: " + (formField.Default ? "checked" : "unchecked"));
                        break;
                    case FieldType.FieldFormTextInput:
                        AppendLine("\tInput format: " + formField.TextInputFormat);
                        AppendLine("\tCurrent contents: " + formField.Result);
                        break;
                }

                // Let the visitor continue visiting other nodes.
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Adds newline char-terminated text to the current output.
            /// </summary>
            private void AppendLine(string text)
            {
                mBuilder.Append(text + '\n');
            }

            /// <summary>
            /// Gets the plain text of the document that was accumulated by the visitor.
            /// </summary>
            public string GetText()
            {
                return mBuilder.ToString();
            }

            private readonly StringBuilder mBuilder;
        }
        //ExEnd

        private void TestFormField(Document doc)
        {
            doc = DocumentHelper.SaveOpen(doc);
            FieldCollection fields = doc.Range.Fields;
            Assert.AreEqual(3, fields.Count);

            Assert.AreEqual(FieldType.FieldFormDropDown, fields[0].Type);
            Assert.AreEqual(" FORMDROPDOWN \u0001", fields[0].GetFieldCode());

            Assert.AreEqual(FieldType.FieldFormCheckBox, fields[1].Type);
            Assert.AreEqual(" FORMCHECKBOX \u0001", fields[1].GetFieldCode());

            Assert.AreEqual(FieldType.FieldFormTextInput, fields[2].Type);
            Assert.AreEqual(" FORMTEXT \u0001", fields[2].GetFieldCode());

            FormFieldCollection formFields = doc.Range.FormFields;
            Assert.AreEqual(3, formFields.Count);

            Assert.AreEqual(FieldType.FieldFormDropDown, formFields[0].Type);
            Assert.AreEqual(new[] { "One", "Two", "Three" }, formFields[0].DropDownItems);
            Assert.True(formFields[0].CalculateOnExit);
            Assert.AreEqual(0, formFields[0].DropDownSelectedIndex);
            Assert.True(formFields[0].Enabled);
            Assert.AreEqual("One", formFields[0].Result);

            Assert.AreEqual(FieldType.FieldFormCheckBox, formFields[1].Type);
            Assert.True(formFields[1].IsCheckBoxExactSize);
            Assert.AreEqual("Right click to check this box", formFields[1].HelpText);
            Assert.True(formFields[1].OwnHelp);
            Assert.AreEqual("Checkbox status text", formFields[1].StatusText);
            Assert.True(formFields[1].OwnStatus);
            Assert.AreEqual(50.0d, formFields[1].CheckBoxSize);
            Assert.False(formFields[1].Checked);
            Assert.False(formFields[1].Default);
            Assert.AreEqual("0", formFields[1].Result);

            Assert.AreEqual(FieldType.FieldFormTextInput, formFields[2].Type);
            Assert.AreEqual("EntryMacro", formFields[2].EntryMacro);
            Assert.AreEqual("ExitMacro", formFields[2].ExitMacro);
            Assert.AreEqual("Regular", formFields[2].TextInputDefault);
            Assert.AreEqual("FIRST CAPITAL", formFields[2].TextInputFormat);
            Assert.AreEqual(TextFormFieldType.Regular, formFields[2].TextInputType);
            Assert.AreEqual(50, formFields[2].MaxLength);
            Assert.AreEqual("This value overrides the one we set during initialization", formFields[2].Result);
        }
    }
}