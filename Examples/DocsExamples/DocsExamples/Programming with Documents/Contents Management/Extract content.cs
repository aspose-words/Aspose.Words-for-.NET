using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Contents_Management
{
    public class ExtractContent : DocsExamplesBase
    {
        [Test]
        public void ExtractContentBetweenBlockLevelNodes()
        {
            //ExStart:ExtractContentBetweenBlockLevelNodes
            //GistId:1f94e59ea4838ffac2f0edf921f67060
            Document doc = new Document(MyDir + "Extract content.docx");

            Paragraph startPara = (Paragraph) doc.LastSection.GetChild(NodeType.Paragraph, 2, true);
            Table endTable = (Table) doc.LastSection.GetChild(NodeType.Table, 0, true);
            // Extract the content between these nodes in the document. Include these markers in the extraction.
            List<Node> extractedNodes = ExtractContentHelper.ExtractContent(startPara, endTable, true, false);

            // Let's reverse the array to make inserting the content back into the document easier.
            extractedNodes.Reverse();
            foreach (Node extractedNode in extractedNodes)
                endTable.ParentNode.InsertAfter(extractedNode, endTable);

            doc.Save(ArtifactsDir + "ExtractContent.ExtractContentBetweenBlockLevelNodes.docx");
            //ExEnd:ExtractContentBetweenBlockLevelNodes
        }

        [Test]
        public void ExtractContentBetweenBookmark()
        {
            //ExStart:ExtractContentBetweenBookmark
            //GistId:1f94e59ea4838ffac2f0edf921f67060
            Document doc = new Document(MyDir + "Extract content.docx");

            Bookmark bookmark = doc.Range.Bookmarks["Bookmark1"];
            BookmarkStart bookmarkStart = bookmark.BookmarkStart;
            BookmarkEnd bookmarkEnd = bookmark.BookmarkEnd;

            // Firstly, extract the content between these nodes, including the bookmark.
            List<Node> extractedNodesInclusive = ExtractContentHelper.ExtractContent(bookmarkStart, bookmarkEnd, true, true);
            
            Document dstDoc = ExtractContentHelper.GenerateDocument(doc, extractedNodesInclusive);
            dstDoc.Save(ArtifactsDir + "ExtractContent.ExtractContentBetweenBookmark.IncludingBookmark.docx");

            // Secondly, extract the content between these nodes this time without including the bookmark.
            List<Node> extractedNodesExclusive = ExtractContentHelper.ExtractContent(bookmarkStart, bookmarkEnd, false, true);
            
            dstDoc = ExtractContentHelper.GenerateDocument(doc, extractedNodesExclusive);
            dstDoc.Save(ArtifactsDir + "ExtractContent.ExtractContentBetweenBookmark.WithoutBookmark.docx");
            //ExEnd:ExtractContentBetweenBookmark
        }

        [Test]
        public void ExtractContentBetweenCommentRange()
        {
            //ExStart:ExtractContentBetweenCommentRange
            //GistId:1f94e59ea4838ffac2f0edf921f67060
            Document doc = new Document(MyDir + "Extract content.docx");

            CommentRangeStart commentStart = (CommentRangeStart) doc.GetChild(NodeType.CommentRangeStart, 0, true);
            CommentRangeEnd commentEnd = (CommentRangeEnd) doc.GetChild(NodeType.CommentRangeEnd, 0, true);

            // Firstly, extract the content between these nodes including the comment as well.
            List<Node> extractedNodesInclusive = ExtractContentHelper.ExtractContent(commentStart, commentEnd, true, true);
            
            Document dstDoc = ExtractContentHelper.GenerateDocument(doc, extractedNodesInclusive);
            dstDoc.Save(ArtifactsDir + "ExtractContent.ExtractContentBetweenCommentRange.IncludingComment.docx");

            // Secondly, extract the content between these nodes without the comment.
            List<Node> extractedNodesExclusive = ExtractContentHelper.ExtractContent(commentStart, commentEnd, false, true);
            
            dstDoc = ExtractContentHelper.GenerateDocument(doc, extractedNodesExclusive);
            dstDoc.Save(ArtifactsDir + "ExtractContent.ExtractContentBetweenCommentRange.WithoutComment.docx");
            //ExEnd:ExtractContentBetweenCommentRange
        }

        [Test]
        public void ExtractContentBetweenParagraphs()
        {
            //ExStart:ExtractContentBetweenParagraphs
            //GistId:1f94e59ea4838ffac2f0edf921f67060
            Document doc = new Document(MyDir + "Extract content.docx");

            Paragraph startPara = (Paragraph) doc.FirstSection.Body.GetChild(NodeType.Paragraph, 6, true);
            Paragraph endPara = (Paragraph) doc.FirstSection.Body.GetChild(NodeType.Paragraph, 10, true);
            // Extract the content between these nodes in the document. Include these markers in the extraction.
            List<Node> extractedNodes = ExtractContentHelper.ExtractContent(startPara, endPara, true, true);

            Document dstDoc = ExtractContentHelper.GenerateDocument(doc, extractedNodes);
            dstDoc.Save(ArtifactsDir + "ExtractContent.ExtractContentBetweenParagraphs.docx");
            //ExEnd:ExtractContentBetweenParagraphs
        }
        
        [Test]
        public void ExtractContentBetweenParagraphStyles()
        {
            //ExStart:ExtractContentBetweenParagraphStyles
            //GistId:1f94e59ea4838ffac2f0edf921f67060
            Document doc = new Document(MyDir + "Extract content.docx");

            // Gather a list of the paragraphs using the respective heading styles.
            List<Paragraph> parasStyleHeading1 = ParagraphsByStyleName(doc, "Heading 1");
            List<Paragraph> parasStyleHeading3 = ParagraphsByStyleName(doc, "Heading 3");

            // Use the first instance of the paragraphs with those styles.
            Node startPara = parasStyleHeading1[0];
            Node endPara = parasStyleHeading3[0];

            // Extract the content between these nodes in the document. Don't include these markers in the extraction.
            List<Node> extractedNodes = ExtractContentHelper.ExtractContent(startPara, endPara, false, true);

            Document dstDoc = ExtractContentHelper.GenerateDocument(doc, extractedNodes);
            dstDoc.Save(ArtifactsDir + "ExtractContent.ExtractContentBetweenParagraphStyles.docx");
            //ExEnd:ExtractContentBetweenParagraphStyles
        }

        //ExStart:ParagraphsByStyleName
        //GistId:1f94e59ea4838ffac2f0edf921f67060
        public List<Paragraph> ParagraphsByStyleName(Document doc, string styleName)
        {
            // Create an array to collect paragraphs of the specified style.
            List<Paragraph> paragraphsWithStyle = new List<Paragraph>();
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            // Look through all paragraphs to find those with the specified style.
            foreach (Paragraph paragraph in paragraphs)
            {
                if (paragraph.ParagraphFormat.Style.Name == styleName)
                    paragraphsWithStyle.Add(paragraph);
            }

            return paragraphsWithStyle;
        }
        //ExEnd:ParagraphsByStyleName

        [Test]
        public void ExtractContentBetweenRuns()
        {
            //ExStart:ExtractContentBetweenRuns
            //GistId:1f94e59ea4838ffac2f0edf921f67060
            Document doc = new Document(MyDir + "Extract content.docx");

            Paragraph para = (Paragraph) doc.GetChild(NodeType.Paragraph, 7, true);
            Run startRun = para.Runs[1];
            Run endRun = para.Runs[4];

            // Extract the content between these nodes in the document. Include these markers in the extraction.
            List<Node> extractedNodes = ExtractContentHelper.ExtractContent(startRun, endRun, true, false);
            foreach (Node extractedNode in extractedNodes)
                Console.WriteLine(extractedNode.ToString(SaveFormat.Text));
            //ExEnd:ExtractContentBetweenRuns
        }

        [Test]
        public void ExtractContentUsingDocumentVisitor()
        {
            //ExStart:ExtractContentUsingDocumentVisitor
            //GistId:1f94e59ea4838ffac2f0edf921f67060
            Document doc = new Document(MyDir + "Extract content.docx");

            ConvertDocToTxt convertToPlainText = new ConvertDocToTxt();
            // Note that every node in the object model has the accept method so the visiting
            // can be executed not only for the whole document, but for any node in the document.
            doc.Accept(convertToPlainText);

            // Once the visiting is complete, we can retrieve the result of the operation,
            // That in this example, has accumulated in the visitor.
            Console.WriteLine(convertToPlainText.GetText());
            //ExEnd:ExtractContentUsingDocumentVisitor
        }

        //ExStart:ConvertDocToTxt
        //GistId:1f94e59ea4838ffac2f0edf921f67060
        /// <summary>
        /// Simple implementation of saving a document in the plain text format. Implemented as a Visitor.
        /// </summary>
        internal class ConvertDocToTxt : DocumentVisitor
        {
            public ConvertDocToTxt()
            {
                mIsSkipText = false;
                mBuilder = new StringBuilder();
            }

            /// <summary>
            /// Gets the plain text of the document that was accumulated by the visitor.
            /// </summary>
            public string GetText()
            {
                return mBuilder.ToString();
            }

            /// <summary>
            /// Called when a Run node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitRun(Run run)
            {
                AppendText(run.Text);
                // Let the visitor continue visiting other nodes.
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FieldStart node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldStart(FieldStart fieldStart)
            {
                // In Microsoft Word, a field code (such as "MERGEFIELD FieldName") follows
                // after a field start character. We want to skip field codes and output field.
                // Result only, therefore we use a flag to suspend the output while inside a field code.
                // Note this is a very simplistic implementation and will not work very well.
                // If you have nested fields in a document.
                mIsSkipText = true;
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FieldSeparator node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldSeparator(FieldSeparator fieldSeparator)
            {
                // Once reached a field separator node, we enable the output because we are
                // now entering the field result nodes.
                mIsSkipText = false;
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FieldEnd node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldEnd(FieldEnd fieldEnd)
            {
                // Make sure we enable the output when reached a field end because some fields
                // do not have field separator and do not have field result.
                mIsSkipText = false;
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when visiting of a Paragraph node is ended in the document.
            /// </summary>
            public override VisitorAction VisitParagraphEnd(Paragraph paragraph)
            {
                // When outputting to plain text we output Cr+Lf characters.
                AppendText(ControlChar.CrLf);
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitBodyStart(Body body)
            {
                // We can detect beginning and end of all composite nodes such as Section, Body, 
                // Table, Paragraph etc and provide custom handling for them.
                mBuilder.Append("*** Body Started ***\r\n");
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitBodyEnd(Body body)
            {
                mBuilder.Append("*** Body Ended ***\r\n");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a HeaderFooter node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitHeaderFooterStart(HeaderFooter headerFooter)
            {
                // Returning this value from a visitor method causes visiting of this
                // Node to stop and move on to visiting the next sibling node
                // The net effect in this example is that the text of headers and footers
                // Is not included in the resulting output
                return VisitorAction.SkipThisNode;
            }

            /// <summary>
            /// Adds text to the current output. Honors the enabled/disabled output flag.
            /// </summary>
            private void AppendText(string text)
            {
                if (!mIsSkipText)
                    mBuilder.Append(text);
            }

            private readonly StringBuilder mBuilder;
            private bool mIsSkipText;
        }
        //ExEnd:ConvertDocToTxt

        [Test]
        public void ExtractContentUsingField()
        {
            //ExStart:ExtractContentUsingField
            //GistId:1f94e59ea4838ffac2f0edf921f67060
            Document doc = new Document(MyDir + "Extract content.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);
            // Pass the first boolean parameter to get the DocumentBuilder to move to the FieldStart of the field.
            // We could also get FieldStarts of a field using GetChildNode method as in the other examples.
            builder.MoveToMergeField("Fullname", false, false);

            // The builder cursor should be positioned at the start of the field.
            FieldStart startField = (FieldStart) builder.CurrentNode;
            Paragraph endPara = (Paragraph) doc.FirstSection.GetChild(NodeType.Paragraph, 5, true);
            // Extract the content between these nodes in the document. Don't include these markers in the extraction.
            List<Node> extractedNodes = ExtractContentHelper.ExtractContent(startField, endPara, false, true);

            Document dstDoc = ExtractContentHelper.GenerateDocument(doc, extractedNodes);
            dstDoc.Save(ArtifactsDir + "ExtractContent.ExtractContentUsingField.docx");
            //ExEnd:ExtractContentUsingField
        }

        [Test]
        public void SimpleExtractText()
        {
            //ExStart:SimpleExtractText
            //GistId:1f94e59ea4838ffac2f0edf921f67060
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.InsertField("MERGEFIELD Field");

            // When converted to text it will not retrieve fields code or special characters,
            // but will still contain some natural formatting characters such as paragraph markers etc. 
            // This is the same as "viewing" the document as if it was opened in a text editor.
            Console.WriteLine("Convert to text result: " + doc.ToString(SaveFormat.Text));
            //ExEnd:SimpleExtractText
        }

        [Test]
        public void ExtractPrintText()
        {
            //ExStart:ExtractText
            //GistId:458eb4fd5bd1de8b06fab4d1ef1acdc6
            Document doc = new Document(MyDir + "Tables.docx");
            
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            // The range text will include control characters such as "\a" for a cell.
            // You can call ToString and pass SaveFormat.Text on the desired node to find the plain text content.
            Console.WriteLine("Contents of the table: ");
            Console.WriteLine(table.Range.Text);
            //ExEnd:ExtractText

            //ExStart:PrintTextRangeRowAndTable
            //GistId:458eb4fd5bd1de8b06fab4d1ef1acdc6
            Console.WriteLine("\nContents of the row: ");
            Console.WriteLine(table.Rows[1].Range.Text);

            Console.WriteLine("\nContents of the cell: ");
            Console.WriteLine(table.LastRow.LastCell.Range.Text);
            //ExEnd:PrintTextRangeRowAndTable
        }

        [Test]
        public void ExtractImages()
        {
            //ExStart:ExtractImages
            //GistId:1f94e59ea4838ffac2f0edf921f67060
            Document doc = new Document(MyDir + "Images.docx");

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageFileName =
                        $"Image.ExportImages.{imageIndex}_{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}";

                    // Note, if you have only an image (not a shape with a text and the image),
                    // you can use shape.GetShapeRenderer().Save(...) method to save the image.
                    shape.ImageData.Save(ArtifactsDir + imageFileName);
                    imageIndex++;
                }
            }
            //ExEnd:ExtractImages
        }

        [Test]
        public void ExtractContentBasedOnStyles()
        {
            //ExStart:ExtractContentBasedOnStyles
            //GistId:a73b495f610523670f0847331ef4d6fc
            Document doc = new Document(MyDir + "Styles.docx");

            List<Paragraph> paragraphs = ParagraphsByStyleName(doc, "Heading 1");
            Console.WriteLine($"Paragraphs with \"Heading 1\" styles ({paragraphs.Count}):");

            foreach (Paragraph paragraph in paragraphs)
                Console.Write(paragraph.ToString(SaveFormat.Text));

            List<Run> runs = RunsByStyleName(doc, "Intense Emphasis");
            Console.WriteLine($"\nRuns with \"Intense Emphasis\" styles ({runs.Count}):");

            foreach (Run run in runs)
                Console.WriteLine(run.Range.Text);
            //ExEnd:ExtractContentBasedOnStyles
        }

        //ExStart:RunsByStyleName
        //GistId:a73b495f610523670f0847331ef4d6fc
        public List<Run> RunsByStyleName(Document doc, string styleName)
        {
            List<Run> runsWithStyle = new List<Run>();
            NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);

            foreach (Run run in runs)
            {
                if (run.Font.Style.Name == styleName)
                    runsWithStyle.Add(run);
            }

            return runsWithStyle;
        }
        //ExEnd:RunsByStyleName
    }
}