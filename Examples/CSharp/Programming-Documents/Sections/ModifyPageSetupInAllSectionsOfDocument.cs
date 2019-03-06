using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Sections
{
    class ModifyPageSetupInAllSectionsOfDocument
    {
        public static void Run()
        {
            // ExStart:ModifyPageSetupInAllSectionsOfDocument
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithSections();
            Document doc = new Document(dataDir + "ModifyPageSetupInAllSections.doc");

            // It is important to understand that a document can contain many sections and each
            // section has its own page setup. In this case we want to modify them all.
            foreach (Section section in doc)
                section.PageSetup.PaperSize = PaperSize.Letter;

            doc.Save(dataDir + "ModifyPageSetupInAllSections_Out.doc");

            // ExEnd:ModifyPageSetupInAllSectionsOfDocument
            Console.WriteLine("\nSections page setup updatd successfully.\nFile saved at " + dataDir);
        }
    }
}
