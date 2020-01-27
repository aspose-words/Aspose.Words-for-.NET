using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class DocumentPageSetup
    {
        public static void Run()
        {
            // ExStart:DocumentPageSetup
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            Document doc = new Document(dataDir + "Document.doc");

            //Set the layout mode for a section allowing to define the document grid behavior
            //Note that the Document Grid tab becomes visible in the Page Setup dialog of MS Word if any Asian language is defined as editing language.
            doc.FirstSection.PageSetup.LayoutMode = SectionLayoutMode.Grid;
            //Set the number of characters per line in the document grid. 
            doc.FirstSection.PageSetup.CharactersPerLine = 30;
            //Set the number of lines per page in the document grid. 
            doc.FirstSection.PageSetup.LinesPerPage = 10;

            dataDir = dataDir + "Document.PageSetup_out.doc";
            doc.Save(dataDir);
            // ExEnd:DocumentPageSetup
            Console.WriteLine("\nPageSetup properties are set successfully.");
        }
    }
}
