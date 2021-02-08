// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Drawing;
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
        public void Create()
        {
            //ExStart
            //ExFor:FormField
            //ExFor:FormField.Result
            //ExFor:FormField.Type
            //ExFor:FormField.Name
            //ExSummary:Shows how to insert a combo box.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Please select a fruit: ");

            // Insert a combo box which will allow a user to choose an option from a collection of strings.
            FormField comboBox = builder.InsertComboBox("MyComboBox", new[] { "Apple", "Banana", "Cherry" }, 0);

            Assert.AreEqual("MyComboBox", comboBox.Name);
            Assert.AreEqual(FieldType.FieldFormDropDown, comboBox.Type);
            Assert.AreEqual("Apple", comboBox.Result);

            // The form field will appear in the form of a "select" html tag.
            doc.Save(ArtifactsDir + "FormFields.Create.html");
            //ExEnd

            doc = new Document(ArtifactsDir + "FormFields.Create.html");
            comboBox = doc.Range.FormFields[0];

            Assert.AreEqual("MyComboBox", comboBox.Name);
            Assert.AreEqual(FieldType.FieldFormDropDown, comboBox.Type);
            Assert.AreEqual("Apple", comboBox.Result);
        }

        [Test]
        public void TextInput()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertTextInput
            //ExSummary:Shows how to insert a text input form field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Please enter text here: ");

            // Insert a text input field, which will allow the user to click it and enter text.
            // Assign some placeholder text that the user may overwrite and pass
            // a maximum text length of 0 to apply no limit on the form field's contents.
            builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "", "Placeholder text", 0);

            // The form field will appear in the form of an "input" html tag, with a type of "text".
            doc.Save(ArtifactsDir + "FormFields.TextInput.html");
            //ExEnd

            doc = new Document(ArtifactsDir + "FormFields.TextInput.html");

            FormField textInput = doc.Range.FormFields[0];

            Assert.AreEqual("TextInput1", textInput.Name);
            Assert.AreEqual(TextFormFieldType.Regular, textInput.TextInputType);
            Assert.AreEqual(string.Empty, textInput.TextInputFormat);
            Assert.AreEqual("Placeholder text", textInput.Result);
            Assert.AreEqual(0, textInput.MaxLength);
        }

        [Test]
        public void DeleteFormField()
        {
            //ExStart
            //ExFor:FormField.RemoveField
            //ExSummary:Shows how to delete a form field.
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

        [Test]
        public void FormFieldFontFormatting()
        {
            //ExStart
            //ExFor:FormField
            //ExSummary:Shows how to formatting the entire FormField, including the field value.
            Document doc = new Document(MyDir + "Form fields.docx");

            FormField formField = doc.Range.FormFields[0];
            formField.Font.Bold = true;
            formField.Font.Size = 24;
            formField.Font.Color = Color.Red;

            formField.Result = "Aspose.FormField";

            doc = DocumentHelper.SaveOpen(doc);
            
            Run formFieldRun = doc.FirstSection.Body.FirstParagraph.Runs[1];

            Assert.AreEqual("Aspose.FormField", formFieldRun.Text);
            Assert.AreEqual(true, formFieldRun.Font.Bold);
            Assert.AreEqual(24, formFieldRun.Font.Size);
            Assert.AreEqual(Color.Red.ToArgb(), formFieldRun.Font.Color.ToArgb());
            //ExEnd
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
        //ExSummary:Shows how insert different kinds of form fields into a document, and process them with using a document visitor implementation.
        [Test] //ExSkip
        public void Visitor()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a document builder to insert a combo box.
            builder.Write("Choose a value from this combo box: ");
            FormField comboBox = builder.InsertComboBox("MyComboBox", new[] { "One", "Two", "Three" }, 0);
            comboBox.CalculateOnExit = true;
            Assert.AreEqual(3, comboBox.DropDownItems.Count);
            Assert.AreEqual(0, comboBox.DropDownSelectedIndex);
            Assert.True(comboBox.Enabled);

            builder.InsertBreak(BreakType.ParagraphBreak);

            // Use a document builder to insert a check box.
            builder.Write("Click this check box to tick/untick it: ");
            FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 50);
            checkBox.IsCheckBoxExactSize = true;
            checkBox.HelpText = "Right click to check this box";
            checkBox.OwnHelp = true;
            checkBox.StatusText = "Checkbox status text";
            checkBox.OwnStatus = true;
            Assert.AreEqual(50.0d, checkBox.CheckBoxSize);
            Assert.False(checkBox.Checked);
            Assert.False(checkBox.Default);

            builder.InsertBreak(BreakType.ParagraphBreak);

            // Use a document builder to insert text input form field.
            builder.Write("Enter text here: ");
            FormField textInput = builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Placeholder text", 50);
            textInput.EntryMacro = "EntryMacro";
            textInput.ExitMacro = "ExitMacro";
            textInput.TextInputDefault = "Regular";
            textInput.TextInputFormat = "FIRST CAPITAL";
            textInput.SetTextInputValue("New placeholder text");
            Assert.AreEqual(TextFormFieldType.Regular, textInput.TextInputType);
            Assert.AreEqual(50, textInput.MaxLength);

            // This collection contains all our form fields.
            FormFieldCollection formFields = doc.Range.FormFields;
            Assert.AreEqual(3, formFields.Count);

            // Fields display our form fields. We can see their field codes by opening this document
            // in Microsoft and pressing Alt + F9. These fields have no switches,
            // and members of the FormField object fully govern their form fields' content.
            Assert.AreEqual(3, doc.Range.Fields.Count);
            Assert.AreEqual(" FORMDROPDOWN \u0001", doc.Range.Fields[0].GetFieldCode());
            Assert.AreEqual(" FORMCHECKBOX \u0001", doc.Range.Fields[1].GetFieldCode());
            Assert.AreEqual(" FORMTEXT \u0001", doc.Range.Fields[2].GetFieldCode());

            // Allow each form field to accept a document visitor.
            FormFieldVisitor formFieldVisitor = new FormFieldVisitor();

            using (IEnumerator<FormField> fieldEnumerator = formFields.GetEnumerator())
                while (fieldEnumerator.MoveNext())
                    fieldEnumerator.Current.Accept(formFieldVisitor);

            Console.WriteLine(formFieldVisitor.GetText());

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "FormFields.Visitor.html");
            TestFormField(doc); //ExSkip
        }

        /// <summary>
        /// Visitor implementation that prints details of form fields that it visits. 
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
                        AppendLine("\tDrop-down items count: " + formField.DropDownItems.Count + ", default selected item index: " + formField.DropDownSelectedIndex);
                        AppendLine("\tDrop-down items: " + string.Join(", ", formField.DropDownItems.ToArray()));
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

            TestUtil.VerifyField(FieldType.FieldFormDropDown, " FORMDROPDOWN \u0001", string.Empty, doc.Range.Fields[0]);
            TestUtil.VerifyField(FieldType.FieldFormCheckBox, " FORMCHECKBOX \u0001", string.Empty, doc.Range.Fields[1]);
            TestUtil.VerifyField(FieldType.FieldFormTextInput, " FORMTEXT \u0001", "New placeholder text", doc.Range.Fields[2]);

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
            Assert.AreEqual("New placeholder text", formFields[2].Result);
        }

        [Test]
        public void DropDownItemCollection()
        {
            //ExStart
            //ExFor:Fields.DropDownItemCollection
            //ExFor:Fields.DropDownItemCollection.Add(String)
            //ExFor:Fields.DropDownItemCollection.Clear
            //ExFor:Fields.DropDownItemCollection.Contains(String)
            //ExFor:Fields.DropDownItemCollection.Count
            //ExFor:Fields.DropDownItemCollection.GetEnumerator
            //ExFor:Fields.DropDownItemCollection.IndexOf(String)
            //ExFor:Fields.DropDownItemCollection.Insert(Int32, String)
            //ExFor:Fields.DropDownItemCollection.Item(Int32)
            //ExFor:Fields.DropDownItemCollection.Remove(String)
            //ExFor:Fields.DropDownItemCollection.RemoveAt(Int32)
            //ExSummary:Shows how to insert a combo box field, and edit the elements in its item collection.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a combo box, and then verify its collection of drop-down items.
            // In Microsoft Word, the user will click the combo box,
            // and then choose one of the items of text in the collection to display.
            string[] items = { "One", "Two", "Three" };
            FormField comboBoxField = builder.InsertComboBox("DropDown", items, 0);
            DropDownItemCollection dropDownItems = comboBoxField.DropDownItems;

            Assert.AreEqual(3, dropDownItems.Count);
            Assert.AreEqual("One", dropDownItems[0]);
            Assert.AreEqual(1, dropDownItems.IndexOf("Two"));
            Assert.IsTrue(dropDownItems.Contains("Three"));

            // There are two ways of adding a new item to an existing collection of drop-down box items.
            // 1 -  Append an item to the end of the collection:
            dropDownItems.Add("Four");

            // 2 -  Insert an item before another item at a specified index:
            dropDownItems.Insert(3, "Three and a half");

            Assert.AreEqual(5, dropDownItems.Count);

            // Iterate over the collection and print every element.
            using (IEnumerator<string> dropDownCollectionEnumerator = dropDownItems.GetEnumerator())
                while (dropDownCollectionEnumerator.MoveNext())
                    Console.WriteLine(dropDownCollectionEnumerator.Current);

            // There are two ways of removing elements from a collection of drop-down items.
            // 1 -  Remove an item with contents equal to the passed string:
            dropDownItems.Remove("Four");

            // 2 -  Remove an item at an index:
            dropDownItems.RemoveAt(3);

            Assert.AreEqual(3, dropDownItems.Count);
            Assert.IsFalse(dropDownItems.Contains("Three and a half"));
            Assert.IsFalse(dropDownItems.Contains("Four"));

            doc.Save(ArtifactsDir + "FormFields.DropDownItemCollection.html");

            // Empty the whole collection of drop-down items.
            dropDownItems.Clear();
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            dropDownItems = doc.Range.FormFields[0].DropDownItems;

            Assert.AreEqual(0, dropDownItems.Count);

            doc = new Document(ArtifactsDir + "FormFields.DropDownItemCollection.html");
            dropDownItems = doc.Range.FormFields[0].DropDownItems;

            Assert.AreEqual(3, dropDownItems.Count);
            Assert.AreEqual("One", dropDownItems[0]);
            Assert.AreEqual("Two", dropDownItems[1]);
            Assert.AreEqual("Three", dropDownItems[2]);
        }
    }
}