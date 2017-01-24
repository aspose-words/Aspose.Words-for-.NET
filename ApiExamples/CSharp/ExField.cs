// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Globalization;
using System.IO;
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
            DocumentBuilder b = new DocumentBuilder(doc);
            b.InsertField("MERGEFIELD Date");

            //ExStart
            //ExId:ChangeCurrentCulture
            //ExSummary:Shows how to change the culture used in formatting fields during update.
            // Store the current culture so it can be set back once mail merge is complete.
            CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;
            // Set to German language so dates and numbers are formatted using this culture during mail merge.
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");

            // Execute mail merge.
            doc.MailMerge.Execute(new string[] {"Date"}, new object[] {DateTime.Now});

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

            // Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document.
            doc.Range.Replace(new Regex("The Beginning"), new InsertTcFieldHandler("Chapter 1", "\\l 1"), false);
        }

        public class InsertTcFieldHandler : IReplacingCallback
        {
            // Store the text and switches to be used for the TC fields.
            private string mFieldText;
            private string mFieldSwitches;

            /// <summary>
            /// The switches to use for each TC field. Can be an empty string or null.
            /// </summary>
            public InsertTcFieldHandler(string switches) : this(string.Empty, switches)
            {
                this.mFieldSwitches = switches;
            }

            /// <summary>
            /// The display text and switches to use for each TC field. Display name can be an empty string or null.
            /// </summary>
            public InsertTcFieldHandler(string text, string switches)
            {
                this.mFieldText = text;
                this.mFieldSwitches = switches;
            }

            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
            {
                // Create a builder to insert the field.
                DocumentBuilder builder = new DocumentBuilder((Document) args.MatchNode.Document);
                // Move to the first node of the match.
                builder.MoveTo(args.MatchNode);

                // If the user specified text to be used in the field as display text then use that, otherwise use the 
                // match string as the display text.
                string insertText;

                if (!string.IsNullOrEmpty(this.mFieldText))
                    insertText = this.mFieldText;
                else
                    insertText = args.Match.Value;

                // Insert the TC field before this node using the specified string as the display text and user defined switches.
                builder.InsertField(string.Format("TC \"{0}\" {1}", insertText, this.mFieldSwitches));

                // We have done what we want so skip replacement.
                return ReplaceAction.Skip;
            }
        }

        //ExEnd

        //Bug: there is no isAfter parameter at BuildAndInsert (exception), need more info from dev
        [Test]
        public void InsertFieldWithFieldBuilder()
        {
            Document doc = new Document();

            //Add some text into the paragraph
            Run run = DocumentHelper.InsertNewRun(doc, " Hello World!", 0);

            FieldArgumentBuilder argumentBuilder = new FieldArgumentBuilder();
            argumentBuilder.AddField(new FieldBuilder(FieldType.FieldMergeField));
            argumentBuilder.AddText("BestField");

            FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FieldIf);
            fieldBuilder.AddArgument(argumentBuilder)
                .AddArgument("=")
                .AddArgument("BestField")
                .AddArgument(10)
                .AddArgument(20.0)
                .AddSwitch("12", "13")
                .BuildAndInsert(run);

            doc.UpdateFields();
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

            Assert.That(() => fieldBuilder.AddArgument(argumentBuilder).AddArgument("=").AddArgument("BestField").AddArgument(10).AddArgument(20.0).BuildAndInsert(run), Throws.TypeOf<ArgumentException>());
        }

        [Test]
        public void BarCodeWord2Pdf()
        {
            Document doc = new Document(MyDir + "BarCode.docx");

            // Set custom barcode generator
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

            doc.Save(MyDir + @"\Artifacts\BarCode.pdf");

            BarCodeReader barCode = BarCodeReaderPdf(MyDir + @"\Artifacts\BarCode.pdf");
            Assert.AreEqual("QR", barCode.GetCodeType().ToString());
        }

        private BarCodeReader BarCodeReaderPdf(string filename)
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
                Console.WriteLine("Codetext found: " + barcodeReader.GetCodeText() + ", Symbology: " + barcodeReader.GetCodeType());
            }

            //close the reader
            barcodeReader.Close();

            return barcodeReader;
        }

        //For assert result of the test you need to open document and check that image are added correct and without truncated inside frame
        [Test]
        public void UpdateFieldIgnoringMergeFormat()
        {
            //ExStart
            //ExFor:FieldIncludePicture.Update(Bool)
            //ExSummary:Shows a way to update a field ignoring the MERGEFORMAT switch
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.PreserveIncludePictureField = true;

            Document doc = new Document(MyDir + "UpdateFieldIgnoringMergeFormat.docx", loadOptions);
            
            foreach (Field field in doc.Range.Fields)
            {
                if (field.Type.Equals(FieldType.FieldIncludePicture))
                {
                    FieldIncludePicture includePicture = (FieldIncludePicture)field;

                    includePicture.SourceFullName = MyDir + "dotnet-logo.png";
                    includePicture.Update(true);
                }
            }
            
            doc.UpdateFields();
            doc.Save(MyDir + @"\Artifacts\UpdateFieldIgnoringMergeFormat.docx");
            //ExEnd
        }

        [Test]
        public void FieldFormat()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Field field = builder.InsertField("MERGEFIELD Date");

            FieldFormat format = field.Format;

            format.DateTimeFormat = "dddd, MMMM dd, yyyy";
            format.NumericFormat = "0.#";
            format.GeneralFormats.Add(GeneralFormat.CharFormat);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            field = doc.Range.Fields[0];
            format = field.Format;

            Assert.AreEqual("0.#", format.NumericFormat);
            Assert.AreEqual("dddd, MMMM dd, yyyy", format.DateTimeFormat);
            Assert.AreEqual(GeneralFormat.CharFormat, format.GeneralFormats[0]);
        }

        [Test]
        public void UpdatePageNumbersInToc()
        {
            Document doc = new Document(MyDir + "Field.UpdateTocPages.docx");

            Node startNode = DocumentHelper.GetParagraph(doc, 2);
            Node endNode = null;

            NodeCollection paragraphCollection = doc.GetChildNodes(NodeType.Paragraph, true);

            foreach (Paragraph para in paragraphCollection)
            {
                // Check all runs in the paragraph for the first page breaks.
                foreach (Run run in para.Runs)
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

            foreach (FieldStart field in fStart)
            {
                FieldType fType = field.FieldType;
                if (fType == FieldType.FieldTOC)
                {
                    Paragraph para = (Paragraph)field.GetAncestor(NodeType.Paragraph);
                    para.Range.UpdateFields();
                    break;
                }
            }

            doc.Save(MyDir + "Field.UpdateTocPages Out.docx");
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
    }
}
