using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document("Section.ModifyPageSetupInAllSections.doc");

            // It is important to understand that a document can contain many sections and each
            // section has its own page setup. In this case we want to modify them all.
            foreach (Section section in doc)
                section.PageSetup.PaperSize = PaperSize.Letter;

            doc.Save("Section.ModifyPageSetupInAllSections Out.doc");
        }
    }
}
