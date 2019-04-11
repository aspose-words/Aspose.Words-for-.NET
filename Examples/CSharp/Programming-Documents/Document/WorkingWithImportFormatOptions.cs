using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class WorkingWithImportFormatOptions
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            // Invokes the InsertDocument method shown above to insert a document at a bookmark.
            SmartStyleBehavior(dataDir);
            KeepSourceNumbering(dataDir);
            IgnoreTextBoxes(dataDir);
        }

        static void SmartStyleBehavior(string dataDir)
        {
            // ExStart:SmartStyleBehavior
            Document srcDoc = new Document(dataDir + "source.docx");
            Document dstDoc = new Document(dataDir + "destination.docx");

            DocumentBuilder builder = new DocumentBuilder(dstDoc);
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.PageBreak);

            ImportFormatOptions options = new ImportFormatOptions();
            options.SmartStyleBehavior = true;
            builder.InsertDocument(srcDoc, ImportFormatMode.UseDestinationStyles, options);
            // ExEnd:SmartStyleBehavior
        }

        static void KeepSourceNumbering(string dataDir)
        {
            // ExStart:KeepSourceNumbering
            Document srcDoc = new Document(dataDir + "source.docx");
            Document dstDoc = new Document(dataDir + "destination.docx");

            ImportFormatOptions importFormatOptions = new ImportFormatOptions();
            // Keep source list formatting when importing numbered paragraphs.
            importFormatOptions.KeepSourceNumbering = true;
            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting, importFormatOptions);

            ParagraphCollection srcParas = srcDoc.FirstSection.Body.Paragraphs;
            foreach (Paragraph srcPara in srcParas)
            {
                Node importedNode = importer.ImportNode(srcPara, false);
                dstDoc.FirstSection.Body.AppendChild(importedNode);
            }

            dstDoc.Save(dataDir + "output.docx");
            // ExEnd:KeepSourceNumbering
        }

        public static void IgnoreTextBoxes(string dataDir)
        {
            // ExStart:IgnoreTextBoxes
            Document srcDoc = new Document(dataDir + "source.docx");
            Document dstDoc = new Document(dataDir + "destination.docx");

            ImportFormatOptions importFormatOptions = new ImportFormatOptions();
            // Keep the source text boxes formatting when importing.
            importFormatOptions.IgnoreTextBoxes = false;
            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting, importFormatOptions);

            ParagraphCollection srcParas = srcDoc.FirstSection.Body.Paragraphs;
            foreach (Paragraph srcPara in srcParas)
            {
                Node importedNode = importer.ImportNode(srcPara, true);
                dstDoc.FirstSection.Body.AppendChild(importedNode);
            }

            dstDoc.Save(dataDir + "output.docx");
            // ExEnd:IgnoreTextBoxes
        }
    }
}
