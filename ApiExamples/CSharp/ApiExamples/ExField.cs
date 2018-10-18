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
using System.Data;
using System.Text;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using Aspose.BarCode;
using Aspose.BarCode.BarCodeRecognition;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;
using NUnit.Framework;
using NUnit.Framework.Constraints;
using System.Drawing;

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
            FieldChar fieldStart = (FieldChar)doc.GetChild(NodeType.FieldStart, 0, true);
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

            FieldStart fieldStart = (FieldStart)doc.GetChild(NodeType.FieldStart, 0, true);

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

            FieldRevNum revNum = (FieldRevNum)doc.Range.Fields[0];
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

            FieldRevNum revNum = (FieldRevNum)doc.Range.Fields[0];
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

            FieldInfo info = (FieldInfo)doc.Range.Fields[0];
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

            FieldInfo info = (FieldInfo)doc.Range.Fields[0];
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
                DocumentBuilder builder = new DocumentBuilder((Document)args.MatchNode.Document);
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
                    FieldIncludePicture includePicture = (FieldIncludePicture)field;

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

            Section newSection = (Section)doc.Sections[0].Clone(true);
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
                    Paragraph para = (Paragraph)field.GetAncestor(NodeType.Paragraph);
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
                    CompositeNode curComposite = (CompositeNode)curNode;
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
            FieldAsk fieldAsk = (FieldAsk)builder.InsertField(FieldType.FieldAsk, true);

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
            FieldAdvance field = (FieldAdvance)builder.InsertField(FieldType.FieldAdvance, true);

            builder.Write("This text is moved up and to the right.");

            Assert.AreEqual(FieldType.FieldAdvance, field.Type);
            Assert.AreEqual(" ADVANCE ", field.GetFieldCode());
            // The second text that the builder added will now be moved
            field.RightOffset = "5";
            field.UpOffset = "5";

            Assert.AreEqual(" ADVANCE  \\r 5 \\u 5", field.GetFieldCode());
            // If we want to move text in the other direction, and try do that by using negative values for the above field members, we will get an error in our document
            // Instead, we need to specify a positive value for the opposite respective field directional variable
            field = (FieldAdvance)builder.InsertField(FieldType.FieldAdvance, true);
            field.DownOffset = "5";
            field.LeftOffset = "100";

            Assert.AreEqual(" ADVANCE  \\d 5 \\l 100", field.GetFieldCode());
            // We are still on one paragraph
            Assert.AreEqual(1, doc.FirstSection.Body.Paragraphs.Count);
            // Since we're setting horizontal and vertical positions next, we need to end the paragraph so the previous line does not get moved with the next one
            builder.Writeln("This text is moved down and to the left, overlapping the previous text.");
            // This time we can also use negative values 
            field = (FieldAdvance)builder.InsertField(FieldType.FieldAdvance, true);
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
            FieldAddressBlock field = (FieldAddressBlock)builder.InsertField(FieldType.FieldAddressBlock, true);

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
        public void FieldCollection()
        {
            //ExStart
            //ExFor:FieldCollection
            //ExFor:FieldCollection.Clear
            //ExFor:FieldCollection.Count
            //ExFor:FieldCollection.GetEnumerator
            //ExFor:FieldCollection.Item(Int32)
            //ExFor:FieldCollection.Remove(Field)
            //ExFor:FieldCollection.Remove(FieldStart)
            //ExFor:FieldCollection.RemoveAt(Int32)
            //ExSummary:Shows how to work with a document's collection of fields.
            // Open a document that has fields
            Document doc = new Document(MyDir + "Document.ContainsFields.docx");

            // Get the collection that contains all the fields in a document
            FieldCollection fields = doc.Range.Fields;
            Assert.AreEqual(5, fields.Count);

            // Iterate over the field collection and print contents and type of every field
            using (IEnumerator<Field> fieldEnumerator = fields.GetEnumerator())
            {
                while (fieldEnumerator.MoveNext())
                {
                    Console.WriteLine("Field found: " + fieldEnumerator.Current.Type);
                    Console.WriteLine("\t{" + fieldEnumerator.Current.GetFieldCode() + "}");
                    Console.WriteLine("\t\"" + fieldEnumerator.Current.Result + "\"");
                }
            }

            // Get a field to remove itself
            fields[0].Remove();
            Assert.AreEqual(4, fields.Count);

            // Remove a field by reference
            Field lastField = fields[3];
            fields.Remove(lastField);
            Assert.AreEqual(3, fields.Count);

            // Remove a field by index
            fields.RemoveAt(2);
            Assert.AreEqual(2, fields.Count);

            // Remove all fields from the document
            fields.Clear();
            Assert.AreEqual(0, fields.Count);
        }

        [Test]
        public void FieldCompare()
        {
            //ExStart
            //ExFor:FieldCompare
            //ExFor:FieldCompare.ComparisonOperator
            //ExFor:FieldCompare.LeftExpression
            //ExFor:FieldCompare.RightExpression
            //ExSummary:Shows how to insert a field that compares expressions.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a compare field using a document builder
            FieldCompare field = (FieldCompare)builder.InsertField(FieldType.FieldCompare, true);

            // Construct a comparison statement
            field.LeftExpression = "3";
            field.ComparisonOperator = "<";
            field.RightExpression = "2";

            // The compare field will print a "0" or "1" depending on the truth of its statement
            // The result of this statement is false, so a "0" will be show up in the document
            Assert.AreEqual(" COMPARE  3 < 2", field.GetFieldCode());

            builder.Writeln();

            // Here a "1" will show up, because the statement is true
            field = (FieldCompare)builder.InsertField(FieldType.FieldCompare, true);
            field.LeftExpression = "5";
            field.ComparisonOperator = "=";
            field.RightExpression = "2 + 3";

            Assert.AreEqual(" COMPARE  5 = \"2 + 3\"", field.GetFieldCode());

            doc.UpdateFields();
            doc.Save(MyDir + @"\Artifacts\Field.Compare.docx");
        }

        [Test]
        public void FieldIf()
        {
            //ExStart
            //ExFor:FieldIf
            //ExFor:FieldIf.ComparisonOperator
            //ExFor:FieldIf.EvaluateCondition
            //ExFor:FieldIf.FalseText
            //ExFor:FieldIf.LeftExpression
            //ExFor:FieldIf.RightExpression
            //ExFor:FieldIf.TrueText
            //ExFor:FieldIfComparisonResult
            //ExSummary:Shows how to insert an if field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Statement 1: ");

            // Use document builder to insert an if field
            FieldIf fieldIf = (FieldIf)builder.InsertField(FieldType.FieldIf, true);

            // The if field will output either the TrueText or FalseText string into the document, depending on the truth of the statement
            // In this case, "0 = 1" is incorrect, so the output will be "False"
            fieldIf.LeftExpression = "0";
            fieldIf.ComparisonOperator = "=";
            fieldIf.RightExpression = "1";
            fieldIf.TrueText = "True";
            fieldIf.FalseText = "False";

            Assert.AreEqual(" IF  0 = 1 True False", fieldIf.GetFieldCode());
            Assert.AreEqual(FieldIfComparisonResult.False, fieldIf.EvaluateCondition());

            // This time, the statement is correct, so the output will be "True"
            builder.Write("\nStatement 2: ");
            fieldIf = (FieldIf)builder.InsertField(FieldType.FieldIf, true);
            fieldIf.LeftExpression = "5";
            fieldIf.ComparisonOperator = "=";
            fieldIf.RightExpression = "2 + 3";
            fieldIf.TrueText = "True";
            fieldIf.FalseText = "False";

            Assert.AreEqual(" IF  5 = \"2 + 3\" True False", fieldIf.GetFieldCode());
            Assert.AreEqual(FieldIfComparisonResult.True, fieldIf.EvaluateCondition());

            doc.UpdateFields();
            doc.Save(MyDir + @"\Artifacts\Field.If.docx");
        }

        [Test]
        public void FieldAutoNum()
        {
            //ExStart
            //ExFor:FieldAutoNum
            //ExFor:FieldAutoNum.SeparatorCharacter
            //ExSummary:Shows how to number paragraphs using autonum fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The two fields we insert here will be automatically numbered 1 and 2
            builder.InsertField(FieldType.FieldAutoNum, true);
            builder.Writeln("\tParagraph 1.");
            builder.InsertField(FieldType.FieldAutoNum, true);
            builder.Writeln("\tParagraph 2.");

            foreach (Field field in doc.Range.Fields)
            {
                if (field.Type == FieldType.FieldAutoNum)
                {
                    // Leaving the FieldAutoNum.SeparatorCharacter field null will set the separator character to '.' by default
                    Assert.IsNull(((FieldAutoNum)field).SeparatorCharacter);

                    // The first character of the string entered here will be used as the separator character
                    ((FieldAutoNum)field).SeparatorCharacter = ":";

                    Assert.AreEqual(" AUTONUM  \\s :", field.GetFieldCode());
                }
            }

            doc.Save(MyDir + @"\Artifacts\Field.AutoNum.docx");
            //ExEnd
        }

        //ExStart
        //ExFor:FieldAutoNumLgl
        //ExFor:FieldAutoNumLgl.RemoveTrailingPeriod
        //ExFor:FieldAutoNumLgl.SeparatorCharacter
        //ExSummary:Shows how to organize a document using autonum legal fields
        [Test] //ExSkip
        public void FieldAutoNumLgl()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // This string will be our paragraph text that
            string loremIpsum =
                "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                "\nUt enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ";

            // In this case our autonum legal field will number our first paragraph as "1."
            InsertNumberedClause(builder, "\tHeading 1", loremIpsum, StyleIdentifier.Heading1);

            // Our heading style number will be 1 again, so this field will keep counting headings at a heading level of 1
            InsertNumberedClause(builder, "\tHeading 2", loremIpsum, StyleIdentifier.Heading1);

            // Our heading style is 2, setting the paragraph numbering depth to 2, setting this field's value to "2.1."
            InsertNumberedClause(builder, "\tHeading 3", loremIpsum, StyleIdentifier.Heading2);

            // Our heading style is 3, so we are going deeper again to "2.1.1."
            InsertNumberedClause(builder, "\tHeading 4", loremIpsum, StyleIdentifier.Heading3);

            // Our heading style is 2, and the next field number at that level is "2.2."
            InsertNumberedClause(builder, "\tHeading 5", loremIpsum, StyleIdentifier.Heading2);

            foreach (Field field in doc.Range.Fields)
            {
                if (field.Type == FieldType.FieldAutoNumLegal)
                {
                    // By default the separator will appear as "." in the document but here it is null
                    Assert.IsNull(((FieldAutoNumLgl)field).SeparatorCharacter);

                    // Change the separator character and remove trailing separators
                    ((FieldAutoNumLgl)field).SeparatorCharacter = ":";
                    ((FieldAutoNumLgl)field).RemoveTrailingPeriod = true;
                    Assert.AreEqual(" AUTONUMLGL  \\s : \\e", field.GetFieldCode());
                }
            }

            doc.Save(MyDir + @"\Artifacts\Field.AutoNumLegal.docx");
        }

        /// <summary>
        /// Get a document builder to insert a clause numbered by an autonum legal field
        /// </summary>
        private void InsertNumberedClause(DocumentBuilder builder, string heading, string contents, StyleIdentifier headingStyle)
        {
            // This legal field will automatically number our clauses, taking heading style level into account
            builder.InsertField(FieldType.FieldAutoNumLegal, true);
            builder.CurrentParagraph.ParagraphFormat.StyleIdentifier = headingStyle;
            builder.Writeln(heading);

            // This text will belong to the auto num legal field above it
            // It will collapse when the arrow next to the corresponding autonum legal field is clicked in MS Word
            builder.CurrentParagraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.BodyText;
            builder.Writeln(contents);
        }
        //ExEnd

        [Test]
        public void FieldAutoNumOut()
        {
            //ExStart
            //ExFor:FieldAutoNumOut
            //ExSummary:Shows how to number paragraphs using autonum outline fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The two fields that we insert here will be numbered 1 and 2
            builder.InsertField(FieldType.FieldAutoNumOutline, true);
            builder.Writeln("\tParagraph 1.");
            builder.InsertField(FieldType.FieldAutoNumOutline, true);
            builder.Writeln("\tParagraph 2.");

            foreach (Field field in doc.Range.Fields)
            {
                if (field.Type == FieldType.FieldAutoNumOutline)
                {
                    Assert.AreEqual(" AUTONUMOUT ", field.GetFieldCode());
                }
            }

            doc.Save(MyDir + @"\Artifacts\Field.AutoNumOut.docx");
            //ExEnd
        }

        [Test]
        public void FieldAutoText()
        {
            //ExStart
            //ExFor:Fields.FieldAutoText
            //ExFor:FieldAutoText.EntryName
            //ExSummary:Shows how to insert an auto text field and reference an auto text building block with it. 
            Document doc = new Document();

            // Create a glossary document and add an AutoText building block
            doc.GlossaryDocument = new GlossaryDocument();
            BuildingBlock buildingBlock = new BuildingBlock(doc.GlossaryDocument);
            buildingBlock.Name = "MyBlock";
            buildingBlock.Gallery = BuildingBlockGallery.AutoText;
            buildingBlock.Category = "General";
            buildingBlock.Description = "MyBlock description";
            buildingBlock.Behavior = BuildingBlockBehavior.Paragraph;
            doc.GlossaryDocument.AppendChild(buildingBlock);

            // Create a source and add it as text content to our building block
            Document buildingBlockSource = new Document();
            DocumentBuilder buildingBlockSourceBuilder = new DocumentBuilder(buildingBlockSource);
            buildingBlockSourceBuilder.Writeln("Hello World!");

            Node buildingBlockContent = doc.GlossaryDocument.ImportNode(buildingBlockSource.FirstSection, true);
            buildingBlock.AppendChild(buildingBlockContent);

            // Create an advance field using document builder
            DocumentBuilder builder = new DocumentBuilder(doc);
            FieldAutoText field = (FieldAutoText)builder.InsertField(FieldType.FieldAutoText, true);

            // Refer to our building block by name
            field.EntryName = "MyBlock";

            // The text content of our building block will be visible in the output
            doc.Save(MyDir + @"\Artifacts\Field.AutoText.dotx");
        }

        //ExStart
        //ExFor:Fields.FieldAutoTextList
        //ExFor:Fields.FieldAutoTextList.EntryName
        //ExFor:Fields.FieldAutoTextList.ListStyle
        //ExFor:Fields.FieldAutoTextList.ScreenTip
        //ExSummary:Shows how to use an AutoTextList field to select from a list of AutoText entries.
        [Test] //ExSkip
        public void FieldAutoTextList()
        {
            Document doc = new Document();

            // Create a glossary document and populate it with auto text entries that our auto text list will let us select from
            doc.GlossaryDocument = new GlossaryDocument();
            AppendAutoTextEntry(doc.GlossaryDocument, "AutoText 1", "Contents of AutoText 1");
            AppendAutoTextEntry(doc.GlossaryDocument, "AutoText 2", "Contents of AutoText 2");
            AppendAutoTextEntry(doc.GlossaryDocument, "AutoText 3", "Contents of AutoText 3");

            // Insert an auto text list using a document builder and change its properties
            DocumentBuilder builder = new DocumentBuilder(doc);
            FieldAutoTextList field = (FieldAutoTextList)builder.InsertField(FieldType.FieldAutoTextList, true);
            field.EntryName = "Right click here to pick an AutoText block"; // This is the text that will be visible in the document
            field.ListStyle = "Heading 1";
            field.ScreenTip = "Hover tip text for AutoTextList goes here";

            Assert.AreEqual("Right click here to pick an AutoText block", field.EntryName); //ExSkip
            Assert.AreEqual("Heading 1", field.ListStyle); //ExSkip
            Assert.AreEqual("Hover tip text for AutoTextList goes here", field.ScreenTip); //ExSkip
            Assert.AreEqual(" AUTOTEXTLIST  \"Right click here to pick an AutoText block\" " +
                            "\\s \"Heading 1\" " +
                            "\\t \"Hover tip text for AutoTextList goes here\"", field.GetFieldCode());

            doc.Save(MyDir + @"\Artifacts\Field.AutoTextList.dotx");
        }

        /// <summary>
        /// Create an AutoText entry and add it to a glossary document
        /// </summary>
        private static void AppendAutoTextEntry(GlossaryDocument glossaryDoc, string name, string contents)
        {
            // Create building block and set it up as an auto text entry
            BuildingBlock buildingBlock = new BuildingBlock(glossaryDoc);
            buildingBlock.Name = name;
            buildingBlock.Gallery = BuildingBlockGallery.AutoText;
            buildingBlock.Category = "General";
            buildingBlock.Behavior = BuildingBlockBehavior.Paragraph;

            // Add content to the building block
            Section section = new Section(glossaryDoc);
            section.AppendChild(new Body(glossaryDoc));
            section.Body.AppendParagraph(contents);
            buildingBlock.AppendChild(section);

            // Add auto text entry to glossary document
            glossaryDoc.AppendChild(buildingBlock);
        }
        //ExEnd

        [Test]
        public void FieldGreetingLine()
        {
            //ExStart
            //ExFor:FieldGreetingLine
            //ExFor:FieldGreetingLine.AlternateText
            //ExFor:FieldGreetingLine.GetFieldNames
            //ExFor:FieldGreetingLine.LanguageId
            //ExFor:FieldGreetingLine.NameFormat
            //ExSummary:Shows how to insert a GREETINGLINE field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a custom greeting field with document builder, and also some content
            FieldGreetingLine fieldGreetingLine = (FieldGreetingLine)builder.InsertField(FieldType.FieldGreetingLine, true);
            builder.Writeln("\n\n\tThis is your custom greeting, created programmatically using Aspose Words!");

            // This array contains strings that correspond to column names in the data table that we will mail merge into our document
            Assert.AreEqual(0, fieldGreetingLine.GetFieldNames().Length);

            // To populate that array, we need to specify a format for our greeting line
            fieldGreetingLine.NameFormat = "<< _BEFORE_ Dear >><< _TITLE0_ >><< _LAST0_ >><< _AFTER_ ,>> ";

            // In this case, our greeting line's field names array now has "Courtesy Title" and "Last Name"
            Assert.AreEqual(2, fieldGreetingLine.GetFieldNames().Length);

            // This string will cover any cases where the data in the data table is incorrect by substituting the malformed name with a string
            fieldGreetingLine.AlternateText = "Sir or Madam";

            // We can set the language ID here too
            fieldGreetingLine.LanguageId = "1033";

            Assert.AreEqual(" GREETINGLINE  \\f \"<< _BEFORE_ Dear >><< _TITLE0_ >><< _LAST0_ >><< _AFTER_ ,>> \" \\e \"Sir or Madam\" \\l 1033", fieldGreetingLine.GetFieldCode());

            // Create a source table for our mail merge that has columns that our greeting line will look for
            System.Data.DataTable table = new System.Data.DataTable("Employees");
            table.Columns.Add("Courtesy Title");
            table.Columns.Add("First Name");
            table.Columns.Add("Last Name");
            table.Rows.Add("Mr.", "John", "Doe");
            table.Rows.Add("Mrs.", "Jane", "Cardholder");
            table.Rows.Add("", "No", "Name"); // This row has an invalid value in the Courtesy Title column, so our greeting will default to the alternate text

            doc.MailMerge.Execute(table);

            doc.UpdateFields();
            doc.Save(MyDir + @"\Artifacts\Field.GreetingLine.docx");
            //ExEnd
        }

        [Test]
        public void FieldListNum()
        {
            //ExStart
            //ExFor:FieldListNum
            //ExFor:FieldListNum.HasListName
            //ExFor:FieldListNum.ListLevel
            //ExFor:FieldListNum.ListName
            //ExFor:FieldListNum.StartingNumber
            //ExSummary:Shows how to number paragraphs with LISTNUM fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a list num field using a document builder
            FieldListNum fieldListNum = (FieldListNum)builder.InsertField(FieldType.FieldListNum, true);

            // Lists start counting at 1 by default, but we can change this number at any time
            // In this case, we'll do a zero-based count
            fieldListNum.StartingNumber = "0";
            builder.Writeln("Paragraph 1");

            // Placing several list num fields in one paragraph increases the list level instead of the current number, in this case resulting in "1)a)i)", list level 3
            builder.InsertField(FieldType.FieldListNum, true);
            builder.InsertField(FieldType.FieldListNum, true);
            builder.InsertField(FieldType.FieldListNum, true);
            builder.Writeln("Paragraph 2");

            // The list level resets with new paragraphs, so to keep counting at a desired list level, we need to set the ListLevel property accordingly
            fieldListNum = (FieldListNum)builder.InsertField(FieldType.FieldListNum, true);
            fieldListNum.ListLevel = "3";
            builder.Writeln("Paragraph 3");

            fieldListNum = (FieldListNum)builder.InsertField(FieldType.FieldListNum, true);

            // Setting this property to this particular value will emulate the AUTONUMOUT field
            fieldListNum.ListName = "OutlineDefault";
            Assert.IsTrue(fieldListNum.HasListName);

            // Start counting from 1
            fieldListNum.StartingNumber = "1";
            builder.Writeln("Paragraph 4");

            // Our fields keep track of the count automatically, but the ListName needs to be set with each new field
            fieldListNum = (FieldListNum)builder.InsertField(FieldType.FieldListNum, true);
            fieldListNum.ListName = "OutlineDefault";
            builder.Writeln("Paragraph 5");

            doc.UpdateFields();
            doc.Save(MyDir + @"\Artifacts\Field.FieldListNum.docx");
            //ExEnd
        }

        public void MergeField()
        {
            //ExStart
            //ExFor:FieldMergeField.#ctor
            //ExFor:FieldMergeField.FieldName
            //ExFor:FieldMergeField.FieldNameNoPrefix
            //ExFor:FieldMergeField.IsMapped
            //ExFor:FieldMergeField.IsVerticalFormatting
            //ExFor:FieldMergeField.TextAfter
            //ExSummary:Shows how to use MERGEFIELD fields to perform a mail merge.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create data source for our merge fields
            System.Data.DataTable table = new System.Data.DataTable("Employees");
            table.Columns.Add("Courtesy Title");
            table.Columns.Add("First Name");
            table.Columns.Add("Last Name");
            table.Rows.Add("Mr.", "John", "Doe");
            table.Rows.Add("Mrs.", "Jane", "Cardholder");

            // Insert a merge field that corresponds to one of our columns and put text before and after it
            FieldMergeField fieldMergeField = (FieldMergeField)builder.InsertField(FieldType.FieldMergeField, true);
            fieldMergeField.FieldName = "Courtesy Title";
            fieldMergeField.IsMapped = true;
            fieldMergeField.IsVerticalFormatting = false;
            fieldMergeField.TextBefore = "Dear ";
            fieldMergeField.TextAfter = " ";

            // Insert another merge field for another column
            // We don't need to use every column to perform a mail merge
            fieldMergeField = (FieldMergeField)builder.InsertField(FieldType.FieldMergeField, true);
            fieldMergeField.FieldName = "Last Name";
            fieldMergeField.TextAfter = ":";

            doc.UpdateFields();
            doc.MailMerge.Execute(table);
            doc.Save(MyDir + @"\Artifacts\Field.MergeField.docx");
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
        //ExFor:FormFieldCollection.Clear
        //ExFor:FormFieldCollection.Count
        //ExFor:FormFieldCollection.GetEnumerator
        //ExFor:FormFieldCollection.Item(Int32)
        //ExFor:FormFieldCollection.Item(String)
        //ExFor:FormFieldCollection.Remove(String)
        //ExFor:FormFieldCollection.RemoveAt(Int32)
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
            Assert.AreEqual(true, comboBox.Enabled);

            builder.Writeln();

            // Use a document builder to insert a check box
            FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 50);
            checkBox.IsCheckBoxExactSize = true;
            checkBox.HelpText = "Right click to check this box";
            checkBox.OwnHelp = true;
            checkBox.StatusText = "Checkbox status text";
            checkBox.OwnStatus = true;
            Assert.AreEqual(50.0d, checkBox.CheckBoxSize);
            Assert.AreEqual(false, checkBox.Checked);
            Assert.AreEqual(false, checkBox.Default);

            builder.Writeln();

            // Use a document builder to insert text input form field
            FormField textInput = builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Your text goes here", 50);
            Assert.AreEqual(3, doc.Range.Fields.Count);
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

            // Iterate over the collection with an enumerator, accepting a visitor with each form field
            FormFieldVisitor formFieldVisitor = new FormFieldVisitor();

            using (IEnumerator<FormField> fieldEnumerator = formFields.GetEnumerator())
            {
                while (fieldEnumerator.MoveNext())
                {
                    fieldEnumerator.Current.Accept(formFieldVisitor);
                }
            }

            Console.WriteLine(formFieldVisitor.GetText());

            doc.UpdateFields();
            doc.Save(MyDir + @"\Artifacts\Field.FormField.docx");
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

        //ExStart
        //ExFor:FieldToc
        //ExFor:FieldToc.BookmarkName
        //ExFor:FieldToc.CustomStyles
        //ExFor:FieldToc.EntrySeparator
        //ExFor:FieldToc.HeadingLevelRange
        //ExFor:FieldToc.HideInWebLayout
        //ExFor:FieldToc.InsertHyperlinks
        //ExFor:FieldToc.PageNumberOmittingLevelRange
        //ExFor:FieldToc.PreserveLineBreaks
        //ExFor:FieldToc.PreserveTabs
        //ExFor:FieldToc.UpdatePageNumbers
        //ExFor:FieldToc.UseParagraphOutlineLevel
        //ExSummary:Shows how to insert a TOC and populate it with entries based on heading styles.
        [Test] //ExSkip
        public void FieldToc()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The table of contents we will insert will accept entries that are only within the scope of this bookmark
            builder.StartBookmark("MyBookmark");

            // Insert a list num field using a document builder
            FieldToc fieldToc = (FieldToc)builder.InsertField(FieldType.FieldTOC, true);

            // Limit possible TOC entries to only those within the bookmark we name here
            fieldToc.BookmarkName = "MyBookmark";

            // Normally paragraphs with a "Heading n" style will be the only ones that will be added to a TOC as entries
            // We can set this attribute to include others, such as the style "Quote" in this case
            fieldToc.CustomStyles = "Quote,Heading 1";

            // Filter out any headings that are outside this range
            fieldToc.HeadingLevelRange = "1-3";

            // Headings in this range won't display their page number in their TOC entry
            fieldToc.PageNumberOmittingLevelRange = "2-5";

            fieldToc.EntrySeparator = "-";
            fieldToc.InsertHyperlinks = true;
            fieldToc.HideInWebLayout = false;
            fieldToc.PreserveLineBreaks = true;
            fieldToc.PreserveTabs = true;
            fieldToc.UseParagraphOutlineLevel = false;

            InsertHeading(builder, "First entry", "Heading 1");
            builder.Writeln("Paragraph text.");
            InsertHeading(builder, "Second entry", "Heading 1");
            InsertHeading(builder, "Third entry", "Quote");

            // These two headings will have the page numbers omitted because they are within the "2-5" range
            InsertHeading(builder, "Fourth entry", "Heading 2");
            InsertHeading(builder, "Fifth entry", "Heading 3");

            // This entry will be omitted because "Heading 4" is outside of the "1-3" range we set earlier
            InsertHeading(builder, "Sixth entry", "Heading 4");

            builder.EndBookmark("MyBookmark");
            builder.Writeln("Paragraph text.");

            // This entry will be omitted because it is outside the bookmark specified by the TOC
            InsertHeading(builder, "Fifth entry", "Heading 1");

            Assert.AreEqual(" TOC  \\b MyBookmark \\t \"Quote,Heading 1\" \\o 1-3 \\n 2-5 \\p - \\h \\x \\w", fieldToc.GetFieldCode());

            fieldToc.UpdatePageNumbers();
            doc.UpdateFields();
            doc.Save(MyDir + @"\Artifacts\Field.FieldTOC.docx");
        }

        /// <summary>
        /// Start a new page and insert a paragraph of a specified style
        /// </summary>
        public void InsertHeading(DocumentBuilder builder, string captionText, string styleName)
        {
            builder.InsertBreak(BreakType.PageBreak);
            string originalStyle = builder.ParagraphFormat.StyleName;
            builder.ParagraphFormat.Style = builder.Document.Styles[styleName];
            builder.Writeln(captionText);
            builder.ParagraphFormat.Style = builder.Document.Styles[originalStyle];
        }
        //ExEnd

        //ExStart
        //ExFor:FieldToc.EntryIdentifier
        //ExFor:FieldToc.EntryLevelRange
        //ExSummary:Shows how to insert a TOC field and filter which TC fields end up as entries.
        [Test] //ExSkip
        public void FieldTocEntryIdentifier()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartBookmark("MyBookmark");

            // Insert a list num field using a document builder
            FieldToc fieldToc = (FieldToc)builder.InsertField(FieldType.FieldTOC, true);
            fieldToc.EntryIdentifier = "A";
            fieldToc.EntryLevelRange = "1-3";

            builder.InsertBreak(BreakType.PageBreak);

            // These two entries will appear in the table
            InsertTocEntry(builder, "TC field 1", "A", "1");
            InsertTocEntry(builder, "TC field 2", "A", "2");

            // These two entries will be omitted because of an incorrect type identifier
            InsertTocEntry(builder, "TC field 3", "B", "1");

            // ...and an out-of-range entry level
            InsertTocEntry(builder, "TC field 4", "A", "5");

            Assert.AreEqual(" TOC  \\f A \\l 1-3", fieldToc.GetFieldCode());

            doc.UpdateFields();
            doc.Save(MyDir + @"\Artifacts\Field.FieldTOC.TC.docx");
        }

        /// <summary>
        /// Insert a table of contents entry via a document builder
        /// </summary>
        public void InsertTocEntry(DocumentBuilder builder, string text, string typeIdentifier, string entryLevel)
        {
            FieldTC fieldTc = (FieldTC)builder.InsertField(FieldType.FieldTOCEntry, true);
            fieldTc.Text = text;
            fieldTc.TypeIdentifier = typeIdentifier;
            fieldTc.EntryLevel = entryLevel;
        }
        //ExEnd

        //ExStart
        //ExFor:FieldToc.TableOfFiguresLabel
        //ExFor:FieldToc.CaptionlessTableOfFiguresLabel
        //ExFor:FieldToc.PrefixedSequenceIdentifier
        //ExFor:FieldToc.SequenceSeparator
        //ExSummary:Insert a TOC field and build the table with SEQ fields.
        [Test] //ExSkip
        public void FieldTocFigure()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a list num field using a document builder
            FieldToc fieldToc = (FieldToc)builder.InsertField(FieldType.FieldTOC, true);
            fieldToc.CaptionlessTableOfFiguresLabel = "Figures";
            fieldToc.PrefixedSequenceIdentifier = "ChapterNum";
            fieldToc.SequenceSeparator = ":";

            // By default, the table of contents 
            fieldToc.TableOfFiguresLabel = "Figure";
            builder.InsertBreak(BreakType.PageBreak);

            // These captions will have a sequence identifier that's the same as the table of figures label in our table of contents,
            // so the table of contents will pick them up
            InsertCaption(builder, "Prefix ", "ChapterNum");
            InsertCaption(builder, " Figure ", "Figure");
            builder.Writeln("\nMy paragraph contents.");
            InsertCaption(builder, "Prefix ", "ChapterNum");
            InsertCaption(builder, " Figure ", "Figure");
            builder.Writeln("\nMy paragraph contents.");

            // This will start a new count and won't be picked up by our table of contents
            InsertCaption(builder, "Figure ", "OtherFigureSequence");
            builder.Writeln("My paragraph contents.");

            Assert.AreEqual(" TOC  \\a Figures \\s ChapterNum \\d : \\c Figure", fieldToc.GetFieldCode());

            fieldToc.UpdatePageNumbers();
            doc.UpdateFields();
            doc.Save(MyDir + @"\Artifacts\Field.FieldTOC.SEQ.docx");
        }

        /// <summary>
        /// Insert a sequence field with preceding text and a specified sequence identifier
        /// </summary>
        public void InsertCaption(DocumentBuilder builder, string precedingText, string sequenceIdentifier)
        {
            builder.Write(precedingText);
            FieldSeq caption = (FieldSeq)builder.InsertField(FieldType.FieldSequence, false);
            caption.SequenceIdentifier = sequenceIdentifier;
        }
        //ExEnd

        [Test]
        [Ignore("WORDSNET-13854")]
        public void FieldCitation()
        {
            //ExStart
            //ExFor:Fields.FieldCitation
            //ExFor:Fields.FieldCitation.AnotherSourceTag
            //ExFor:Fields.FieldCitation.FormatLanguageId
            //ExFor:Fields.FieldCitation.PageNumber
            //ExFor:Fields.FieldCitation.Prefix
            //ExFor:Fields.FieldCitation.SourceTag
            //ExFor:Fields.FieldCitation.Suffix
            //ExFor:Fields.FieldCitation.SuppressAuthor
            //ExFor:Fields.FieldCitation.SuppressTitle
            //ExFor:Fields.FieldCitation.SuppressYear
            //ExFor:Fields.FieldCitation.VolumeNumber
            //ExSummary:Shows how to insert a citation field and edit its properties.
            // Open a document that has bibliographical sources
            Document doc = new Document(MyDir + "Document.HasBibliography.docx");

            // Add text that we can cite
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Text to be cited with one source.");

            // Create a citation field using the document builder
            FieldCitation field = (FieldCitation)builder.InsertField(FieldType.FieldCitation, true);

            // A simple citation can have just the page number and author's name
            field.SourceTag = "Book1"; // We refer to sources using their tag names
            field.PageNumber = "85";
            field.SuppressAuthor = false;
            field.SuppressTitle = true;
            field.SuppressYear = true;

            Assert.AreEqual(" CITATION  Book1 \\p 85 \\t \\y", field.GetFieldCode());

            // We can make a more detailed citation and make it cite 2 sources
            builder.Write("Text to be cited with two sources.");
            field = (FieldCitation)builder.InsertField(FieldType.FieldCitation, true);
            field.SourceTag = "Book1";
            field.AnotherSourceTag = "Book2";
            field.FormatLanguageId = "en-US";
            field.PageNumber = "19";
            field.Prefix = "Prefix ";
            field.Suffix = " Suffix";
            field.SuppressAuthor = false;
            field.SuppressTitle = false;
            field.SuppressYear = false;
            field.VolumeNumber = "VII";

            Assert.AreEqual(" CITATION  Book1 \\m Book2 \\l en-US \\p 19 \\f \"Prefix \" \\s \" Suffix\" \\v VII", field.GetFieldCode());

            doc.UpdateFields();
            doc.Save(MyDir + @"\Artifacts\Field.Citation.docx");
            //ExEnd
        }
        
        [Test]
        public void FieldData()
        {
            //ExStart
            //ExFor:FieldData.#ctor
            //ExSummary:Shows how to insert a data field into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a document builder to insert a data field
            FieldData field = (FieldData)builder.InsertField(FieldType.FieldData, true);
            Assert.AreEqual(" DATA ", field.GetFieldCode());
            //ExEnd
        }
        
        [Test]
        public void FieldInclude()
        {
            //ExStart
            //ExFor:FieldInclude.#ctor
            //ExFor:FieldInclude.BookmarkName
            //ExFor:FieldInclude.LockFields
            //ExFor:FieldInclude.SourceFullName
            //ExFor:FieldInclude.TextConverter
            //ExSummary:Shows how to create an INCLUDE field and set its properties.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add an INCLUDE field with document builder and import a portion of the document defined by a bookmark
            FieldInclude fieldInclude = (FieldInclude)builder.InsertField(FieldType.FieldInclude, true);
            fieldInclude.SourceFullName = MyDir + "Field.Include.Source.docx";
            fieldInclude.BookmarkName = "Source_paragraph_2";
            fieldInclude.LockFields = false;
            fieldInclude.TextConverter = "Microsoft Word";

            doc.UpdateFields();
            doc.Save(MyDir + @"\Artifacts\Field.Include.docx");
            //ExEnd
        }

        [Test]
        [Ignore("WORDSNET-13854")]
        public void FieldDatabase()
        {
            //ExStart
            //ExFor:FieldDatabase
            //ExFor:FieldDatabase.Connection
            //ExFor:FieldDatabase.FileName
            //ExFor:FieldDatabase.FirstRecord
            //ExFor:FieldDatabase.FormatAttributes
            //ExFor:FieldDatabase.InsertHeadings
            //ExFor:FieldDatabase.InsertOnceOnMailMerge
            //ExFor:FieldDatabase.LastRecord
            //ExFor:FieldDatabase.Query
            //ExFor:FieldDatabase.TableFormat
            //ExSummary:Shows how to extract data from a database and insert it as a field into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a document builder to insert a database field
            FieldDatabase field = (FieldDatabase)builder.InsertField(FieldType.FieldDatabase, true);

            // Create a simple query that extracts one table from the database
            field.FileName = MyDir + @"Database\Northwind.mdb";
            field.Connection = "DSN=MS Access Databases";
            field.Query = "SELECT * FROM [Products]";

            // Insert another database field
            field = (FieldDatabase)builder.InsertField(FieldType.FieldDatabase, true);
            field.FileName = MyDir + @"Database\Northwind.mdb";
            field.Connection = "DSN=MS Access Databases";

            // This query will sort all the products by their gross sales in descending order
            field.Query =
                "SELECT [Products].ProductName, FORMAT(SUM([Order Details].UnitPrice * (1 - [Order Details].Discount) * [Order Details].Quantity), 'Currency') AS GrossSales " +
                "FROM([Products] " +
                "LEFT JOIN[Order Details] ON[Products].[ProductID] = [Order Details].[ProductID]) " +
                "GROUP BY[Products].ProductName " +
                "ORDER BY SUM([Order Details].UnitPrice* (1 - [Order Details].Discount) * [Order Details].Quantity) DESC";

            // You can use these variables instead of a LIMIT clause, to simplify your query
            // In this case we are taking the first 10 values of the result of our query
            field.FirstRecord = "1";
            field.LastRecord = "10";

            // The number we put here is the index of the format we want to use for our table
            // The list of table formats is in the "Table AutoFormat..." menu we find in MS Word when we create a data table field
            // Index "10" corresponds to the "Colorful 3" format
            field.TableFormat = "10";

            // This attribute decides which elements of the table format we picked above we incorporate into our table
            // The number we use is a sum of a combination of values corresponding to which elements we choose
            // 63 represents borders (1) + shading (2) + font (4) + colour (8) + autofit (16) + heading rows (32)
            field.FormatAttributes = "63";

            field.InsertHeadings = true;
            field.InsertOnceOnMailMerge = true;

            doc.UpdateFields();
            doc.Save(MyDir + @"\Artifacts\Field.Database.docx");
            //ExEnd
        }
        
        [Test]
        public void FieldIncludePicture()
        {
            //ExStart
            //ExFor:FieldIncludePicture.#ctor
            //ExFor:FieldIncludePicture.GraphicFilter
            //ExFor:FieldIncludePicture.IsLinked
            //ExFor:FieldIncludePicture.ResizeHorizontally
            //ExFor:FieldIncludePicture.ResizeVertically
            //ExFor:FieldIncludePicture.SourceFullName
            //ExSummary:Shows how to create an INCLUDEPICTURE field and set its properties.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            FieldIncludePicture fieldIncludePicture = (FieldIncludePicture)builder.InsertField(FieldType.FieldIncludePicture, true);
            fieldIncludePicture.SourceFullName = MyDir + "Images/Watermark.png";

            // Apply, in this case, the PNG32.FLT filter
            fieldIncludePicture.GraphicFilter = "PNG32";
            fieldIncludePicture.IsLinked = true;
            fieldIncludePicture.ResizeHorizontally = true;
            fieldIncludePicture.ResizeVertically = true;

            doc.UpdateFields();
            doc.Save(MyDir + @"\Artifacts\Field.IncludePicture.docx");
            //ExEnd
        }

        //ExStart
        //ExFor:FieldIncludeText.#ctor
        //ExFor:FieldIncludeText.BookmarkName
        //ExFor:FieldIncludeText.Encoding
        //ExFor:FieldIncludeText.LockFields
        //ExFor:FieldIncludeText.MimeType
        //ExFor:FieldIncludeText.NamespaceMappings
        //ExFor:FieldIncludeText.SourceFullName
        //ExFor:FieldIncludeText.TextConverter
        //ExFor:FieldIncludeText.XPath
        //ExFor:FieldIncludeText.XslTransformation
        //ExSummary:Shows how to create an INCLUDETEXT field and set its properties.
        [Test] //ExSkip
        [Ignore("WORDSNET-17543")] //ExSkip
        public void FieldIncludeText()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a document builder to insert an include text field and perform an XSL transformation on an XML document
            FieldIncludeText fieldIncludeText = CreateFieldIncludeText(builder, MyDir + "Field.IncludeText.Source.xml", false, "text/xml", "XML", "ISO-8859-1");
            fieldIncludeText.XslTransformation = MyDir + "Field.IncludeText.Source.xsl";

            builder.Writeln();

            // Use a document builder to insert an include text field and use an XPath to take specific elements
            fieldIncludeText = CreateFieldIncludeText(builder, MyDir + "Field.IncludeText.Source.xml", false, "text/xml", "XML", "ISO-8859-1");
            fieldIncludeText.NamespaceMappings = "xmlns:n='myNamespace'";
            fieldIncludeText.XPath = "/catalog/cd/title";

            doc.Save(MyDir + @"\Artifacts\Field.IncludeText.docx");
        }

        /// <summary>
        /// Use a document builder to insert an INCLUDETEXT field and set its properties
        /// </summary>
        public FieldIncludeText CreateFieldIncludeText(DocumentBuilder builder, string sourceFullName, bool lockFields, string mimeType, string textConverter, string encoding)
        {
            FieldIncludeText fieldIncludeText = (FieldIncludeText)builder.InsertField(FieldType.FieldIncludeText, true);
            fieldIncludeText.SourceFullName = sourceFullName;
            fieldIncludeText.LockFields = lockFields;
            fieldIncludeText.MimeType = mimeType;
            fieldIncludeText.TextConverter = textConverter;
            fieldIncludeText.Encoding = encoding;

            return fieldIncludeText;
        }
        //ExEnd
        
        [Test] 
        [Ignore("WORDSNET-17545")]
        public void FieldHyperlink()
        {
            //ExStart
            //ExFor:FieldHyperlink.#ctor
            //ExFor:FieldHyperlink.Address
            //ExFor:FieldHyperlink.IsImageMap
            //ExFor:FieldHyperlink.OpenInNewWindow
            //ExFor:FieldHyperlink.ScreenTip
            //ExFor:FieldHyperlink.SubAddress
            //ExFor:FieldHyperlink.Target
            //ExSummary:Shows how to insert HYPERLINK fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a hyperlink with a document builder
            FieldHyperlink fieldHyperlink = (FieldHyperlink)builder.InsertField(FieldType.FieldHyperlink, true);

            // When link is clicked, open a document and place the cursor on the bookmarked location
            fieldHyperlink.Address = MyDir + "Field.HyperlinkDestination.docx";
            fieldHyperlink.SubAddress = "My_Bookmark";
            fieldHyperlink.ScreenTip = "Open " + fieldHyperlink.Address + " on bookmark " + fieldHyperlink.SubAddress + " in a new window";

            builder.Writeln();

            // Open html file at a specific frame
            fieldHyperlink = (FieldHyperlink)builder.InsertField(FieldType.FieldHyperlink, true);
            fieldHyperlink.Address = MyDir + "Field.HyperlinkDestination.html";
            fieldHyperlink.ScreenTip = "Open " + fieldHyperlink.Address;
            fieldHyperlink.Target = "iframe_3";
            fieldHyperlink.OpenInNewWindow = true;
            fieldHyperlink.IsImageMap = false;

            doc.UpdateFields();
            doc.Save(MyDir + @"\Artifacts\Field.Hyperlink.docx");
            //ExEnd
        }

        [Test]
        public void FieldBarcode()
        {
            //ExStart
            //ExFor:FieldBarcode
            //ExFor:FieldBarcode.FacingIdentificationMark
            //ExFor:FieldBarcode.IsBookmark
            //ExFor:FieldBarcode.IsUSPostalAddress
            //ExFor:FieldBarcode.PostalAddress
            //ExSummary:Shows how to insert a BARCODE field and set its properties. 
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a document builder to insert a bookmark with a US postal code in it
            builder.StartBookmark("BarcodeBookmark");
            builder.Writeln("96801");
            builder.EndBookmark("BarcodeBookmark");

            builder.Writeln();

            // Reference a US postal code directly
            FieldBarcode fieldBarcode = (FieldBarcode)builder.InsertField(FieldType.FieldBarcode, true);
            fieldBarcode.FacingIdentificationMark = "C";
            fieldBarcode.PostalAddress = "96801";
            fieldBarcode.IsUSPostalAddress = true;

            builder.Writeln();

            // Reference a US postal code from a bookmark
            fieldBarcode = (FieldBarcode)builder.InsertField(FieldType.FieldBarcode, true);
            fieldBarcode.PostalAddress = "BarcodeBookmark";
            fieldBarcode.IsBookmark = true;

            doc.UpdateFields();
            doc.Save(MyDir + @"\Artifacts\Field.USAddressBarcode.docx");
        }

        [Test]
        public void FieldDisplayBarcode()
        {
            //ExStart
            //ExFor:FieldDisplayBarcode
            //ExFor:FieldDisplayBarcode.AddStartStopChar
            //ExFor:FieldDisplayBarcode.BackgroundColor
            //ExFor:FieldDisplayBarcode.BarcodeType
            //ExFor:FieldDisplayBarcode.BarcodeValue
            //ExFor:FieldDisplayBarcode.CaseCodeStyle
            //ExFor:FieldDisplayBarcode.DisplayText
            //ExFor:FieldDisplayBarcode.ErrorCorrectionLevel
            //ExFor:FieldDisplayBarcode.FixCheckDigit
            //ExFor:FieldDisplayBarcode.ForegroundColor
            //ExFor:FieldDisplayBarcode.PosCodeStyle
            //ExFor:FieldDisplayBarcode.ScalingFactor
            //ExFor:FieldDisplayBarcode.SymbolHeight
            //ExFor:FieldDisplayBarcode.SymbolRotation
            //ExSummary:Shows how to insert a DISPLAYBARCODE field and set its properties. 
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            FieldDisplayBarcode fieldDisplayBarcode = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

            // Insert a QR code
            fieldDisplayBarcode.BarcodeType = "QR";
            fieldDisplayBarcode.BarcodeValue = "ABC123";
            fieldDisplayBarcode.BackgroundColor = "0xF8BD69";
            fieldDisplayBarcode.ForegroundColor = "0xB5413B";
            fieldDisplayBarcode.ErrorCorrectionLevel = "3";
            fieldDisplayBarcode.ScalingFactor = "250";
            fieldDisplayBarcode.SymbolHeight = "1000";
            fieldDisplayBarcode.SymbolRotation = "0";

            builder.Writeln();

            // insert a EAN13 barcode
            fieldDisplayBarcode = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            fieldDisplayBarcode.BarcodeType = "EAN13";
            fieldDisplayBarcode.BarcodeValue = "501234567890";         
            fieldDisplayBarcode.DisplayText = true;
            fieldDisplayBarcode.PosCodeStyle = "CASE";
            fieldDisplayBarcode.FixCheckDigit = true;

            builder.Writeln();

            // insert a CODE39 barcode
            fieldDisplayBarcode = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            fieldDisplayBarcode.BarcodeType = "CODE39";
            fieldDisplayBarcode.BarcodeValue = "12345ABCDE";
            fieldDisplayBarcode.AddStartStopChar = true;

            builder.Writeln();

            // insert a ITF14 barcode
            fieldDisplayBarcode = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            fieldDisplayBarcode.BarcodeType = "ITF14";
            fieldDisplayBarcode.BarcodeValue = "09312345678907";
            fieldDisplayBarcode.CaseCodeStyle = "STD";

            doc.UpdateFields();
            doc.Save(MyDir + @"\Artifacts\Field.DisplayBarcode.docx");
            //ExEnd
        }

        //ExStart
        //ExFor:BarcodeParameters
        //ExFor:BarcodeParameters.AddStartStopChar
        //ExFor:BarcodeParameters.BackgroundColor
        //ExFor:BarcodeParameters.BarcodeType
        //ExFor:BarcodeParameters.BarcodeValue
        //ExFor:BarcodeParameters.CaseCodeStyle
        //ExFor:BarcodeParameters.DisplayText
        //ExFor:BarcodeParameters.ErrorCorrectionLevel
        //ExFor:BarcodeParameters.FacingIdentificationMark
        //ExFor:BarcodeParameters.FixCheckDigit
        //ExFor:BarcodeParameters.ForegroundColor
        //ExFor:BarcodeParameters.IsBookmark
        //ExFor:BarcodeParameters.IsUSPostalAddress
        //ExFor:BarcodeParameters.PosCodeStyle
        //ExFor:BarcodeParameters.PostalAddress
        //ExFor:BarcodeParameters.ScalingFactor
        //ExFor:BarcodeParameters.SymbolHeight
        //ExFor:BarcodeParameters.SymbolRotation
        //ExFor:IBarcodeGenerator
        //ExFor:IBarcodeGenerator.GetBarcodeImage(BarcodeParameters)
        //ExFor:IBarcodeGenerator.GetOldBarcodeImage(BarcodeParameters)
        //ExSummary:Shows how to create barcode images using a barcode generator.
        [Test] //ExSkip
        public void BarcodeGenerator()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Assert.IsNull(doc.FieldOptions.BarcodeGenerator);

            // Barcodes generated in this way will be images, and we can use a custom IBarcodeGenerator implementation to generate them
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

            // Configure barcode parameters for a QR barcode
            BarcodeParameters barcodeParameters = new BarcodeParameters();
            barcodeParameters.BarcodeType = "QR";
            barcodeParameters.BarcodeValue = "ABC123";
            barcodeParameters.BackgroundColor = "0xF8BD69";
            barcodeParameters.ForegroundColor = "0xB5413B";
            barcodeParameters.ErrorCorrectionLevel = "3";
            barcodeParameters.ScalingFactor = "250";
            barcodeParameters.SymbolHeight = "1000";
            barcodeParameters.SymbolRotation = "0";

            // Save the generated barcode image to the file system
            Image img = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters);
            img.Save(MyDir + @"\Artifacts\Field.BarcodeGenerator.QR.jpg");

            // Insert the image into the document
            builder.InsertImage(img);

            // Configure barcode parameters for a EAN13 barcode
            barcodeParameters = new BarcodeParameters();
            barcodeParameters.BarcodeType = "EAN13";
            barcodeParameters.BarcodeValue = "501234567890";
            barcodeParameters.DisplayText = true;
            barcodeParameters.PosCodeStyle = "CASE";
            barcodeParameters.FixCheckDigit = true;

            img = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters);
            img.Save(MyDir + @"\Artifacts\Field.BarcodeGenerator.EAN13.jpg");
            builder.InsertImage(img);

            // Configure barcode parameters for a CODE39 barcode
            barcodeParameters = new BarcodeParameters();
            barcodeParameters.BarcodeType = "CODE39";
            barcodeParameters.BarcodeValue = "12345ABCDE";
            barcodeParameters.AddStartStopChar = true;

            img = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters);
            img.Save(MyDir + @"\Artifacts\Field.BarcodeGenerator.CODE39.jpg");
            builder.InsertImage(img);

            // Configure barcode parameters for an ITF14 barcode
            barcodeParameters = new BarcodeParameters();
            barcodeParameters.BarcodeType = "ITF14";
            barcodeParameters.BarcodeValue = "09312345678907";
            barcodeParameters.CaseCodeStyle = "STD";

            img = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters);
            img.Save(MyDir + @"\Artifacts\Field.BarcodeGenerator.ITF14.jpg");
            builder.InsertImage(img);

            doc.Save(MyDir + @"\Artifacts\Field.BarcodeGenerator.docx");
        }
        //ExEnd
    }
}