using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class DocumentBuilderMovingCursor
    {
        public static void Run()
        {
            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            CursorPosition(dataDir);
            MoveToNode(dataDir);
            MoveToDocumentStartEnd(dataDir);
            MoveToSection(dataDir);
            HeadersAndFooters(dataDir);
            MoveToParagraph(dataDir);
            MoveToTableCell(dataDir);
            MoveToBookmark(dataDir);
            MoveToBookmarkEnd(dataDir);
            MoveToMergeField(dataDir);

        }
        public static void CursorPosition(string dataDir)
        {
            //ExStart:DocumentBuilderCursorPosition
            // Shows how to access the current node in a document builder.
            Document doc = new Document(dataDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            Node curNode = builder.CurrentNode;
            Paragraph curParagraph = builder.CurrentParagraph;
            //ExEnd:DocumentBuilderCursorPosition
            Console.WriteLine("\nCursor move to paragraph: " + curParagraph.GetText());
        }
        public static void MoveToNode(string dataDir)
        {
            //ExStart:DocumentBuilderMoveToNode
            Document doc = new Document(dataDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveTo(doc.FirstSection.Body.LastParagraph);
            //ExEnd:DocumentBuilderMoveToNode   
            Console.WriteLine("\nCursor move to required node.");
        }
        public static void MoveToDocumentStartEnd(string dataDir)
        {
            //ExStart:DocumentBuilderMoveToDocumentStartEnd
            Document doc = new Document(dataDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToDocumentEnd();
            Console.WriteLine("\nThis is the end of the document.");

            builder.MoveToDocumentStart();
            Console.WriteLine("\nThis is the beginning of the document.");
            //ExEnd:DocumentBuilderMoveToDocumentStartEnd            
        }
        public static void MoveToSection(string dataDir)
        {
            //ExStart:DocumentBuilderMoveToSection
            Document doc = new Document(dataDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Parameters are 0-index. Moves to third section.
            builder.MoveToSection(2);
            builder.Writeln("This is the 3rd section.");
            //ExEnd:DocumentBuilderMoveToSection               
        }
        public static void HeadersAndFooters(string dataDir)
        {
            //ExStart:DocumentBuilderHeadersAndFooters
            // Create a blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Specify that we want headers and footers different for first, even and odd pages.
            builder.PageSetup.DifferentFirstPageHeaderFooter = true;
            builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

            // Create the headers.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.Write("Header First");
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
            builder.Write("Header Even");
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Header Odd");

            // Create three pages in the document.
            builder.MoveToSection(0);
            builder.Writeln("Page1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page3");

            dataDir = dataDir + "DocumentBuilder.HeadersAndFooters_out_.doc";
            doc.Save(dataDir);
            //ExEnd:DocumentBuilderHeadersAndFooters   
            Console.WriteLine("\nHeaders and footers created successfully using DocumentBuilder.\nFile saved at " + dataDir);
        }
        public static void MoveToParagraph(string dataDir)
        {
            //ExStart:DocumentBuilderMoveToParagraph
            Document doc = new Document(dataDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Parameters are 0-index. Moves to third paragraph.
            builder.MoveToParagraph(2, 0);
            builder.Writeln("This is the 3rd paragraph.");
            //ExEnd:DocumentBuilderMoveToParagraph               
        }
        public static void MoveToTableCell(string dataDir)
        {
            //ExStart:DocumentBuilderMoveToTableCell
            Document doc = new Document(dataDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // All parameters are 0-index. Moves to the 2nd table, 3rd row, 5th cell.
            builder.MoveToCell(1, 2, 4, 0);
            builder.Writeln("Hello World!");
            //ExEnd:DocumentBuilderMoveToTableCell               
        }
        public static void MoveToBookmark(string dataDir)
        {
            //ExStart:DocumentBuilderMoveToBookmark
            Document doc = new Document(dataDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToBookmark("CoolBookmark");
            builder.Writeln("This is a very cool bookmark.");
            //ExEnd:DocumentBuilderMoveToBookmark               
        }
        public static void MoveToBookmarkEnd(string dataDir)
        {
            //ExStart:DocumentBuilderMoveToBookmarkEnd
            Document doc = new Document(dataDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToBookmark("CoolBookmark", false, true);
            builder.Writeln("This is a very cool bookmark.");
            //ExEnd:DocumentBuilderMoveToBookmarkEnd              
        }
        public static void MoveToMergeField(string dataDir)
        {
            //ExStart:DocumentBuilderMoveToMergeField
            Document doc = new Document(dataDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToMergeField("NiceMergeField");
            builder.Writeln("This is a very nice merge field.");
            //ExEnd:DocumentBuilderMoveToMergeField              
        }     
    }
}
