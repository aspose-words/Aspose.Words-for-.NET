// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using Aspose.BarCode.BarCodeRecognition;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExField : ApiExampleBase
    {
        [Test]
        public void UpdateToc()
        {
            Document doc = new Document();

            //ExStart
            //ExId:UpdateTOC
            //ExSummary:Shows how to completely rebuild TOC fields in the document by invoking field update.
            doc.UpdateFields();
            //ExEnd
        }

        [Test]
        public void GetFieldType()
        {
            Document doc = new Document(MyDir + "Document.TableOfContents.doc");

            //ExStart
            //ExFor:FieldType
            //ExFor:FieldChar
            //ExFor:FieldChar.FieldType
            //ExSummary:Shows how to find the type of field that is represented by a node which is derived from FieldChar.
            FieldChar fieldStart = (FieldChar) doc.GetChild(NodeType.FieldStart, 0, true);
            FieldType type = fieldStart.FieldType;
            //ExEnd
        }

        [Test]
        public void GetFieldFromDocument()
        {
            //ExStart
            //ExFor:FieldChar.GetField
            //ExFor:Field.IsLocked
            //ExId:GetField
            //ExSummary:Demonstrates how to retrieve the field class from an existing FieldStart node in the document.
            Document doc = new Document(MyDir + "Document.TableOfContents.doc");

            FieldStart fieldStart = (FieldStart) doc.GetChild(NodeType.FieldStart, 0, true);

            // Retrieve the facade object which represents the field in the document.
            Field field = fieldStart.GetField();

            Console.WriteLine("Field code:" + field.GetFieldCode());
            Console.WriteLine("Field result: " + field.Result);
            Console.WriteLine("Is locked: " + field.IsLocked);

            // This updates only this field in the document.
            field.Update();
            //ExEnd
        }

        [Test]
        public void CreateRevNumFieldWithFieldBuilder()
        {
            //ExStart
            //ExFor:FieldBuilder.#ctor(FieldType)
            //ExFor:FieldBuilder.BuildAndInsert(Inline)
            //ExSummary:Builds and inserts a field into the document before the specified inline node
            Document doc = new Document();
            Run run = DocumentHelper.InsertNewRun(doc, " Hello World!", 0);

            FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FieldRevisionNum);
            fieldBuilder.BuildAndInsert(run);

            doc.UpdateFields();
            //ExEnd
            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            FieldRevNum revNum = (FieldRevNum) doc.Range.Fields[0];
            Assert.NotNull(revNum);
        }

        [Test]
        public void CreateRevNumFieldByDocumentBuilder()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertField("REVNUM MERGEFORMAT");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            FieldRevNum revNum = (FieldRevNum) doc.Range.Fields[0];
            Assert.NotNull(revNum);
        }

        [Test]
        public void CreateInfoFieldWithFieldBuilder()
        {
            Document doc = new Document();
            Run run = DocumentHelper.InsertNewRun(doc, " Hello World!", 0);

            FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FieldInfo);
            fieldBuilder.BuildAndInsert(run);

            doc.UpdateFields();

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            FieldInfo info = (FieldInfo) doc.Range.Fields[0];
            Assert.NotNull(info);
        }

        [Test]
        public void CreateInfoFieldWithDocumentBuilder()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertField("INFO MERGEFORMAT");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            FieldInfo info = (FieldInfo) doc.Range.Fields[0];
            Assert.NotNull(info);
        }

        [Test]
        public void GetFieldFromFieldCollection()
        {
            //ExStart
            //ExId:GetFieldFromFieldCollection
            //ExSummary:Demonstrates how to retrieve a field using the range of a node.
            Document doc = new Document(MyDir + "Document.TableOfContents.doc");

            Field field = doc.Range.Fields[0];

            // This should be the first field in the document - a TOC field.
            Console.WriteLine(field.Type);
            //ExEnd
        }

        [Test]
        public void InsertTcField()
        {
            //ExStart
            //ExId:InsertTCField
            //ExSummary:Shows how to insert a TC field into the document using DocumentBuilder.
            // Create a blank document.
            Document doc = new Document();

            // Create a document builder to insert content with.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a TC field at the current document builder position.
            builder.InsertField("TC \"Entry Text\" \\f t");
            //ExEnd
        }

        [Test]
        public void ChangeLocale()
        {
            // Create a blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertField("MERGEFIELD Date");

            //ExStart
            //ExId:ChangeCurrentCulture
            //ExSummary:Shows how to change the culture used in formatting fields during update.
            // Store the current culture so it can be set back once mail merge is complete.
            CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;
            // Set to German language so dates and numbers are formatted using this culture during mail merge.
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");

            // Execute mail merge.
            doc.MailMerge.Execute(new[] { "Date" }, new object[] { DateTime.Now });

            // Restore the original culture.
            Thread.CurrentThread.CurrentCulture = currentCulture;
            //ExEnd

            doc.Save(MyDir + @"\Artifacts\Field.ChangeLocale.doc");
        }

        [Test]
        public void RemoveTocFromDocument()
        {
            //ExStart
            //ExFor:CompositeNode.GetChildNodes(NodeType, Boolean)
            //ExId:RemoveTableOfContents
            //ExSummary:Demonstrates how to remove a specified TOC from a document.
            // Open a document which contains a TOC.
            Document doc = new Document(MyDir + "Document.TableOfContents.doc");

            // Remove the first TOC from the document.
            Field tocField = doc.Range.Fields[0];
            tocField.Remove();

            // Save the output.
            doc.Save(MyDir + @"\Artifacts\Document.TableOfContentsRemoveTOC.doc");
            //ExEnd
        }

        [Test]
        //ExStart
        //ExId:TCFieldsRangeReplace
        //ExSummary:Shows how to find and insert a TC field at text in a document.
        public void InsertTcFieldsAtText()
        {
            Document doc = new Document();

            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new InsertTcFieldHandler("Chapter 1", "\\l 1");

            // Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document.
            doc.Range.Replace(new Regex("The Beginning"), "", options);
        }

        private class InsertTcFieldHandler : IReplacingCallback
        {
            // Store the text and switches to be used for the TC fields.
            private readonly String mFieldText;
            private readonly String mFieldSwitches;

            /// <summary>
            /// The display text and switches to use for each TC field. Display name can be an empty String or null.
            /// </summary>
            public InsertTcFieldHandler(String text, String switches)
            {
                mFieldText = text;
                mFieldSwitches = switches;
            }

            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
            {
                // Create a builder to insert the field.
                DocumentBuilder builder = new DocumentBuilder((Document) args.MatchNode.Document);
                // Move to the first node of the match.
                builder.MoveTo(args.MatchNode);

                // If the user specified text to be used in the field as display text then use that, otherwise use the 
                // match String as the display text.
                String insertText;

                if (!String.IsNullOrEmpty(mFieldText))
                    insertText = mFieldText;
                else
                    insertText = args.Match.Value;

                // Insert the TC field before this node using the specified String as the display text and user defined switches.
                builder.InsertField(String.Format("TC \"{0}\" {1}", insertText, mFieldSwitches));

                // We have done what we want so skip replacement.
                return ReplaceAction.Skip;
            }
        }
        //ExEnd

        [Test]
        [Ignore("WORDSNET-16037")]
        public void InsertAndUpdateDirtyField()
        {
            //ExStart
            //ExFor:Field.IsDirty
            //ExSummary:Shows how to use special property for updating field result
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Field fieldToc = builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            fieldToc.IsDirty = true;
            //ExEnd

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            // Assert that field model is correct
            Assert.IsTrue(doc.Range.Fields[0].IsDirty);

            LoadOptions loadOptions = new LoadOptions { UpdateDirtyFields = false };

            doc = new Document(dstStream, loadOptions);
            Field tocField = doc.Range.Fields[0];

            // Assert that isDirty saves 
            Assert.IsTrue(tocField.IsDirty);
        }

        [Test]
        public void InsertFieldWithFieldBuilder()
        {
            //ExStart
            //ExFor:FieldArgumentBuilder.#ctor
            //ExFor:FieldArgumentBuilder.AddField(FieldBuilder)
            //ExFor:FieldArgumentBuilder.AddText(String)
            //ExFor:FieldBuilder.#ctor
            //ExFor:FieldBuilder.AddArgument(FieldArgumentBuilder)
            //ExFor:FieldBuilder.AddArgument(String)
            //ExFor:FieldBuilder.AddArgument(Int32)
            //ExFor:FieldBuilder.AddArgument(Double)
            //ExFor:FieldBuilder.AddSwitch(String, String)
            //ExSummary:Inserts a field into a document using field builder constructor
            Document doc = new Document();

            //Add text into the paragraph
            Paragraph para = doc.FirstSection.Body.Paragraphs[0];
            Run run = new Run(doc) { Text = " Hello World!" };
            para.AppendChild(run);

            FieldArgumentBuilder argumentBuilder = new FieldArgumentBuilder();
            argumentBuilder.AddField(new FieldBuilder(FieldType.FieldMergeField));
            argumentBuilder.AddText("BestField");

            FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FieldIf);
            fieldBuilder.AddArgument(argumentBuilder).AddArgument("=").AddArgument("BestField").AddArgument(10)
                .AddArgument(20.0).AddSwitch("12", "13").BuildAndInsert(run);

            doc.UpdateFields();
            //ExEnd
        }

        [Test]
        public void InsertFieldWithFieldBuilderException()
        {
            Document doc = new Document();

            //Add some text into the paragraph
            Run run = DocumentHelper.InsertNewRun(doc, " Hello World!", 0);

            FieldArgumentBuilder argumentBuilder = new FieldArgumentBuilder();
            argumentBuilder.AddField(new FieldBuilder(FieldType.FieldMergeField));
            argumentBuilder.AddNode(run);
            argumentBuilder.AddText("Text argument builder");

            FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FieldIncludeText);

            Assert.That(
                () => fieldBuilder.AddArgument(argumentBuilder).AddArgument("=").AddArgument("BestField")
                    .AddArgument(10).AddArgument(20.0).BuildAndInsert(run), Throws.TypeOf<ArgumentException>());
        }
#if !(NETSTANDARD2_0 || __MOBILE__)
        [Test]
        public void BarCodeWord2Pdf()
        {
            Document doc = new Document(MyDir + "Field.BarCode.docx");

            // Set custom barcode generator
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

            doc.Save(MyDir + @"\Artifacts\Field.BarCode.pdf");

            BarCodeReader barCode = BarCodeReaderPdf(MyDir + @"\Artifacts\Field.BarCode.pdf");
            Assert.AreEqual("QR", barCode.GetCodeType().ToString());
        }

        private BarCodeReader BarCodeReaderPdf(String filename)
        {
            //Set license for Aspose.BarCode
            Aspose.BarCode.License licenceBarCode = new Aspose.BarCode.License();
            licenceBarCode.SetLicense(@"X:\awnet\TestData\Licenses\Aspose.Total.lic");

            //bind the pdf document
            Aspose.Pdf.Facades.PdfExtractor pdfExtractor = new Aspose.Pdf.Facades.PdfExtractor();
            pdfExtractor.BindPdf(filename);

            //set page range for image extraction
            pdfExtractor.StartPage = 1;
            pdfExtractor.EndPage = 1;

            pdfExtractor.ExtractImage();

            //save image to stream
            MemoryStream imageStream = new MemoryStream();
            pdfExtractor.GetNextImage(imageStream);
            imageStream.Position = 0;

            //recognize the barcode from the image stream above
            BarCodeReader barcodeReader = new BarCodeReader(imageStream, DecodeType.QR);
            while (barcodeReader.Read())
            {
                Console.WriteLine("Codetext found: " + barcodeReader.GetCodeText() + ", Symbology: " +
                                  barcodeReader.GetCodeType());
            }

            //close the reader
            barcodeReader.Close();

            return barcodeReader;
        }
#endif
        //For assert result of the test you need to open document and check that image are added correct and without truncated inside frame
        [Test]
        public void UpdateFieldIgnoringMergeFormat()
        {
            //ExStart
            //ExFor:Field.Update(bool)
            //ExSummary:Shows a way to update a field ignoring the MERGEFORMAT switch
            LoadOptions loadOptions = new LoadOptions { PreserveIncludePictureField = true };

            Document doc = new Document(MyDir + "Field.UpdateFieldIgnoringMergeFormat.docx", loadOptions);

            foreach (Field field in doc.Range.Fields)
            {
                if (field.Type.Equals(FieldType.FieldIncludePicture))
                {
                    FieldIncludePicture includePicture = (FieldIncludePicture) field;

                    includePicture.SourceFullName = MyDir + @"\Images\dotnet-logo.png";
                    includePicture.Update(true);
                }
            }

            doc.UpdateFields();
            doc.Save(MyDir + @"\Artifacts\Field.UpdateFieldIgnoringMergeFormat.docx");
            //ExEnd
        }

        [Test]
        public void FieldFormat()
        {
            //ExStart
            //ExFor:Field.Format
            //ExFor:FieldFormat
            //ExFor:FieldFormat.DateTimeFormat
            //ExFor:FieldFormat.NumericFormat
            //ExFor:FieldFormat.GeneralFormats
            //ExFor:GeneralFormat
            //ExFor:GeneralFormatCollection.Add(GeneralFormat)
            //ExSummary:Shows how to formatting fields
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Field field = builder.InsertField("MERGEFIELD Date");

            FieldFormat format = field.Format;

            format.DateTimeFormat = "dddd, MMMM dd, yyyy";
            format.NumericFormat = "0.#";
            format.GeneralFormats.Add(GeneralFormat.CharFormat);
            //ExEnd

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            field = doc.Range.Fields[0];
            format = field.Format;

            Assert.AreEqual("0.#", format.NumericFormat);
            Assert.AreEqual("dddd, MMMM dd, yyyy", format.DateTimeFormat);
            Assert.AreEqual(GeneralFormat.CharFormat, format.GeneralFormats[0]);
        }

        [Test]
        public void UnlinkAllFieldsInDocument()
        {
            //ExStart
            //ExFor:Document.UnlinkFields
            //ExSummary:Shows how to unlink all fields in the document
            Document doc = new Document(MyDir + "Field.UnlinkFields.docx");

            doc.UnlinkFields();
            //ExEnd

            String paraWithFields = DocumentHelper.GetParagraphText(doc, 0);
            Assert.AreEqual("Fields.Docx   Элементы указателя не найдены.     1.\r", paraWithFields);
        }

        [Test]
        public void UnlinkAllFieldsInRange()
        {
            //ExStart
            //ExFor:Range.UnlinkFields
            //ExSummary:Shows how to unlink all fields in range
            Document doc = new Document(MyDir + "Field.UnlinkFields.docx");

            Section newSection = (Section) doc.Sections[0].Clone(true);
            doc.Sections.Add(newSection);

            doc.Sections[1].Range.UnlinkFields();
            //ExEnd

            String secWithFields = DocumentHelper.GetSectionText(doc, 1);
            Assert.AreEqual(
                "Fields.Docx   Элементы указателя не найдены.     3.\rОшибка! Не указана последовательность.    Fields.Docx   Элементы указателя не найдены.     4.\r\r\r\r\r\f",
                secWithFields);
        }

        [Test]
        public void UnlinkSingleField()
        {
            //ExStart
            //ExFor:Field.Unlink
            //ExSummary:Shows how to unlink specific field
            Document doc = new Document(MyDir + "Field.UnlinkFields.docx");
            doc.Range.Fields[1].Unlink();
            //ExEnd

            String paraWithFields = DocumentHelper.GetParagraphText(doc, 0);
            Assert.AreEqual(
                "\u0013 FILENAME  \\* Caps  \\* MERGEFORMAT \u0014Fields.Docx\u0015   Элементы указателя не найдены.     \u0013 LISTNUM  LegalDefault \u0015\r",
                paraWithFields);
        }

        [Test]
        public void UpdatePageNumbersInToc()
        {
            Document doc = new Document(MyDir + "Field.UpdateTocPages.docx");

            Node startNode = DocumentHelper.GetParagraph(doc, 2);
            Node endNode = null;

            NodeCollection paragraphCollection = doc.GetChildNodes(NodeType.Paragraph, true);

            foreach (Paragraph para in paragraphCollection.OfType<Paragraph>())
            {
                // Check all runs in the paragraph for the first page breaks.
                foreach (Run run in para.Runs.OfType<Run>())
                {
                    if (run.Text.Contains(ControlChar.PageBreak))
                    {
                        endNode = run;
                        break;
                    }
                }
            }

            if (startNode != null && endNode != null)
            {
                RemoveSequence(startNode, endNode);

                startNode.Remove();
                endNode.Remove();
            }

            NodeCollection fStart = doc.GetChildNodes(NodeType.FieldStart, true);

            foreach (FieldStart field in fStart.OfType<FieldStart>())
            {
                FieldType fType = field.FieldType;
                if (fType == FieldType.FieldTOC)
                {
                    Paragraph para = (Paragraph) field.GetAncestor(NodeType.Paragraph);
                    para.Range.UpdateFields();
                    break;
                }
            }

            doc.Save(MyDir + @"\Artifacts\Field.UpdateTocPages.docx");
        }

        private void RemoveSequence(Node start, Node end)
        {
            Node curNode = start.NextPreOrder(start.Document);
            while (curNode != null && !curNode.Equals(end))
            {
                //Move to next node
                Node nextNode = curNode.NextPreOrder(start.Document);

                //Check whether current contains end node
                if (curNode.IsComposite)
                {
                    CompositeNode curComposite = (CompositeNode) curNode;
                    if (!curComposite.GetChildNodes(NodeType.Any, true).Contains(end) &&
                        !curComposite.GetChildNodes(NodeType.Any, true).Contains(start))
                    {
                        nextNode = curNode.NextSibling;
                        curNode.Remove();
                    }
                }
                else
                {
                    curNode.Remove();
                }

                curNode = nextNode;
            }
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
            //ExSummary:Shows how to insert a combo box field and manipulate the elements in its item collection.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a document builder to create and populate a combo box
            string[] items = { "One", "Two", "Three" };
            FormField comboBoxField = builder.InsertComboBox("DropDown", items, 0);

            // Get the list of drop down items
            DropDownItemCollection dropDownItems = comboBoxField.DropDownItems;

            Assert.AreEqual(3, dropDownItems.Count);
            Assert.AreEqual("One", dropDownItems[0]);
            Assert.AreEqual(1, dropDownItems.IndexOf("Two"));
            Assert.IsTrue(dropDownItems.Contains("Three"));

            // We can add an item to the end of the collection or insert it at a desired index
            dropDownItems.Add("Four");
            dropDownItems.Insert(3, "Three and a half");
            Assert.AreEqual(5, dropDownItems.Count);

            // Iterate over the collection and print every element
            using (IEnumerator<string> dropDownCollectionEnumerator = dropDownItems.GetEnumerator())
            {
                while (dropDownCollectionEnumerator.MoveNext())
                {
                    string currentItem = dropDownCollectionEnumerator.Current;
                    Console.WriteLine(currentItem);
                }
            }

            // We can remove elements in the same way we added them
            dropDownItems.Remove("Four");
            dropDownItems.RemoveAt(3);
            Assert.IsFalse(dropDownItems.Contains("Three and a half"));
            Assert.IsFalse(dropDownItems.Contains("Four"));

            doc.Save(MyDir + @"\Artifacts\Fields.DropDownItems.docx");

            // Empty the collection
            dropDownItems.Clear();
            Assert.AreEqual(0, dropDownItems.Count);
        }

        [Test]
        public void FieldAsk()
        {
            //ExStart
            //ExFor:Fields.FieldAsk
            //ExFor:Fields.FieldAsk.BookmarkName
            //ExFor:Fields.FieldAsk.DefaultResponse
            //ExFor:Fields.FieldAsk.PromptOnceOnMailMerge
            //ExFor:Fields.FieldAsk.PromptText
            //ExSummary:Shows how to create an ASK field and set its properties.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // We can use a document builder to create our field
            FieldAsk fieldAsk = (FieldAsk) builder.InsertField(FieldType.FieldAsk, true);

            // The initial state of our ask field is empty
            Assert.AreEqual(" ASK ", fieldAsk.GetFieldCode());

            // Add functionality to our field
            fieldAsk.BookmarkName = "MyAskField";
            fieldAsk.PromptText = "Please provide a response for this ASK field";
            fieldAsk.DefaultResponse = "This is the default response.";
            fieldAsk.PromptOnceOnMailMerge = true;

            // The attributes we changed are now incorporated into the field code
            Assert.AreEqual(
                " ASK  MyAskField \"Please provide a response for this ASK field\" \\d \"This is the default response.\" \\o",
                fieldAsk.GetFieldCode());
            //ExEnd

            Assert.AreEqual("MyAskField", fieldAsk.BookmarkName);
            Assert.AreEqual("Please provide a response for this ASK field", fieldAsk.PromptText);
            Assert.AreEqual("This is the default response.", fieldAsk.DefaultResponse);
            Assert.AreEqual(true, fieldAsk.PromptOnceOnMailMerge);
        }

        [Test]
        public void FieldAdvance()
        {
            //ExStart
            //ExFor:Fields.FieldAdvance
            //ExFor:Fields.FieldAdvance.DownOffset
            //ExFor:Fields.FieldAdvance.HorizontalPosition
            //ExFor:Fields.FieldAdvance.LeftOffset
            //ExFor:Fields.FieldAdvance.RightOffset
            //ExFor:Fields.FieldAdvance.UpOffset
            //ExFor:Fields.FieldAdvance.VerticalPosition
            //ExSummary:Shows how to insert an advance field and edit its properties. 
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("This text is in its normal place.");
            // Create an advance field using document builder
            FieldAdvance field = (FieldAdvance) builder.InsertField(FieldType.FieldAdvance, true);

            builder.Write("This text is moved up and to the right.");

            Assert.AreEqual(FieldType.FieldAdvance, field.Type);
            Assert.AreEqual(" ADVANCE ", field.GetFieldCode());
            // The second text that the builder added will now be moved
            field.RightOffset = "5";
            field.UpOffset = "5";

            Assert.AreEqual(" ADVANCE  \\r 5 \\u 5", field.GetFieldCode());
            // If we want to move text in the other direction, and try do that by using negative values for the above field members, we will get an error in our document
            // Instead, we need to specify a positive value for the opposite respective field directional variable
            field = (FieldAdvance) builder.InsertField(FieldType.FieldAdvance, true);
            field.DownOffset = "5";
            field.LeftOffset = "100";

            Assert.AreEqual(" ADVANCE  \\d 5 \\l 100", field.GetFieldCode());
            // We are still on one paragraph
            Assert.AreEqual(1, doc.FirstSection.Body.Paragraphs.Count);
            // Since we're setting horizontal and vertical positions next, we need to end the paragraph so the previous line does not get moved with the next one
            builder.Writeln("This text is moved down and to the left, overlapping the previous text.");
            // This time we can also use negative values 
            field = (FieldAdvance) builder.InsertField(FieldType.FieldAdvance, true);
            field.HorizontalPosition = "-100";
            field.VerticalPosition = "200";

            Assert.AreEqual(" ADVANCE  \\x -100 \\y 200", field.GetFieldCode());

            builder.Write("This text is in a custom position.");

            doc.Save(MyDir + @"\Artifacts\Field.Advance.docx");
            //ExEnd
        }

        [Test]
        public void FieldAddressBlock()
        {
            //ExStart
            //ExFor:Fields.FieldAddressBlock.ExcludedCountryOrRegionName
            //ExFor:Fields.FieldAddressBlock.FormatAddressOnCountryOrRegion
            //ExFor:Fields.FieldAddressBlock.IncludeCountryOrRegionName
            //ExFor:Fields.FieldAddressBlock.LanguageId
            //ExFor:Fields.FieldAddressBlock.NameAndAddressFormat
            //ExSummary:Shows how to build a field address block.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            // Use a document builder to insert a field address block
            FieldAddressBlock field = (FieldAddressBlock) builder.InsertField(FieldType.FieldAddressBlock, true);
            // Initially our field is an empty address block field with null attributes
            Assert.AreEqual(" ADDRESSBLOCK ", field.GetFieldCode());
            // Setting this to "2" will cause all countries/regions to be included, unless it is the one specified in the ExcludedCountryOrRegionName attribute
            field.IncludeCountryOrRegionName = "2";
            field.FormatAddressOnCountryOrRegion = true;
            field.ExcludedCountryOrRegionName = "United States";
            // Specify our own name and address format
            field.NameAndAddressFormat = "<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>";
            // By default, the language ID will be set to that of the first character of the document
            // In this case we will specify it to be English
            field.LanguageId = "1033";
            // Our field code has changed according to the attribute values that we set
            Assert.AreEqual(
                " ADDRESSBLOCK  \\c 2 \\d \\e \"United States\" \\f \"<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>\" \\l 1033",
                field.GetFieldCode());
            //ExEnd
            Assert.AreEqual("2", field.IncludeCountryOrRegionName);
            Assert.AreEqual(true, field.FormatAddressOnCountryOrRegion);
            Assert.AreEqual("United States", field.ExcludedCountryOrRegionName);
            Assert.AreEqual("<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>",
                field.NameAndAddressFormat);
            Assert.AreEqual("1033", field.LanguageId);
        }

        [Test]
        [Ignore("InsertAsHtml issue (3rd field)")]
        public void FieldLink()
        {
            //ExStart
            //ExFor:FieldLink
            //ExFor:FieldLink.AutoUpdate
            //ExFor:FieldLink.FormatUpdateType
            //ExFor:FieldLink.InsertAsBitmap
            //ExFor:FieldLink.InsertAsHtml
            //ExFor:FieldLink.InsertAsPicture
            //ExFor:FieldLink.InsertAsRtf
            //ExFor:FieldLink.InsertAsText
            //ExFor:FieldLink.InsertAsUnicode
            //ExFor:FieldLink.IsLinked
            //ExFor:FieldLink.ProgId
            //ExFor:FieldLink.SourceFullName
            //ExFor:FieldLink.SourceItem
            //ExSummary:Shows how to create link fields with various sources and presentation types.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            // Use a document builder to insert a field link
            // Here we will insert a spreadsheet as a bitmap image
            FieldLink field = (FieldLink)builder.InsertField(FieldType.FieldLink, true);
            field.InsertAsBitmap = true;
            field.AutoUpdate = true;
            field.ProgId = "Excel.Sheet.8";
            field.SourceFullName = MyDir + "MySpreadsheet.xlsx";
            // Setting this field to "4" will keep the source format
            field.FormatUpdateType = "4";
            builder.Writeln();

            // Inserting one cell from a spreadsheet as text
            field = (FieldLink)builder.InsertField(FieldType.FieldLink, true);
            field.InsertAsText = true;
            field.AutoUpdate = true;
            field.ProgId = "Excel.Sheet.8";
            // Take only one cell from the source spreadsheet
            field.SourceItem = "Sheet1!R2C2";
            field.SourceFullName = MyDir + "MySpreadsheet.xlsx";
            builder.Writeln();

            // Inserting a word document as HTML format text
            field = (FieldLink)builder.InsertField(FieldType.FieldLink, true);
            field.InsertAsHtml = true;
            field.AutoUpdate = true;
            field.ProgId = "Word.Document.8";
            field.SourceFullName = MyDir + "Document.doc";
            builder.Writeln();

            // Inserting a document as a rtf
            field = (FieldLink)builder.InsertField(FieldType.FieldLink, true);
            field.InsertAsRtf = true;
            field.AutoUpdate = true;
            field.ProgId = "Word.Document.8";
            field.SourceFullName = MyDir + "Document.doc";
            builder.Writeln();

            // Inserting a document as unicode text
            field = (FieldLink)builder.InsertField(FieldType.FieldLink, true);
            field.InsertAsUnicode = true;
            field.AutoUpdate = true;
            field.ProgId = "Word.Document.8";
            field.SourceFullName = MyDir + "Document.doc";
            builder.Writeln();

            // Insert an image
            field = (FieldLink)builder.InsertField(FieldType.FieldLink, true);
            field.InsertAsPicture = true;
            field.AutoUpdate = true;
            field.ProgId = "Paint.Picture";
            field.SourceFullName = MyDir + "Images/Test_1024_768.bmp";
            // Setting this to true will not store the data in the document, reducing file size
            field.IsLinked = true;
            builder.Writeln();

            // You will be prompted to let the fields update when you open this document, give it a few seconds to do so
            doc.UpdateFields();
            doc.Save(MyDir + @"\Artifacts\Field.Link.docx");
            //ExEnd
        }
    }
}