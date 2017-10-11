using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class WorkingWithFootnote
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            SetFootNoteColumns(dataDir);
            SetFootnoteOptions(dataDir);
            SetEndnoteOptions(dataDir);


        }

        private static void SetFootNoteColumns(string dataDir)
        {
            // ExStart:SetFootNoteColumns
            Document doc = new Document(dataDir + "TestFile.docx");
            
            //Specify the number of columns with which the footnotes area is formatted. 
            doc.FootnoteOptions.Columns = 3;
            dataDir = dataDir + "TestFile_Out.doc";

            // Save the document to disk.
            doc.Save(dataDir);
            // ExEnd:SetFootNoteColumns      
            Console.WriteLine("\nFootnote number of columns set successfully.\nFile saved at " + dataDir);
        }

        private static void SetFootnoteAndEndNotePosition(string dataDir)
        {
            // ExStart:SetFootnoteAndEndNotePosition
            Document doc = new Document(dataDir + "TestFile.docx");

            //Set footnote and endnode position.
            doc.FootnoteOptions.Position = FootnotePosition.BeneathText;
            doc.EndnoteOptions.Position = EndnotePosition.EndOfSection;
            dataDir = dataDir + "TestFile_Out.doc";

            // Save the document to disk.
            doc.Save(dataDir);
            // ExEnd:SetFootnoteAndEndNotePosition      
            Console.WriteLine("\nFootnote number of columns set successfully.\nFile saved at " + dataDir);
        }
        private static void SetEndnoteOptions(string dataDir)
        {
            // ExStart:SetEndnoteOptions
            Document doc = new Document(dataDir + "TestFile.docx");

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Some text");

            builder.InsertFootnote(FootnoteType.Endnote, "Eootnote text.");

            EndnoteOptions option = doc.EndnoteOptions;
            option.RestartRule = FootnoteNumberingRule.RestartPage;
            option.Position = EndnotePosition.EndOfSection;

            dataDir = dataDir + "TestFile_Out.doc";

            // Save the document to disk.
            doc.Save(dataDir);
            // ExEnd:SetEndnoteOptions      
            Console.WriteLine("\nEootnote is inserted at the end of section successfully.\nFile saved at " + dataDir);
        }
    }
}
