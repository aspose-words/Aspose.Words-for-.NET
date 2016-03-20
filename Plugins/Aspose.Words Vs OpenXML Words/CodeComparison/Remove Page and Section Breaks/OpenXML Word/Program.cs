using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML_Word
{
    class Program
    {
        static void Main(string[] args)
        {
            RemovePageBreaks("Test.docx");
        }
        static void RemovePageBreaks(string filename)
        {

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filename, true))
            {

                MainDocumentPart mainPart = myDoc.MainDocumentPart;

                List<Break> breaks = mainPart.Document.Descendants<Break>().ToList();

                foreach (Break b in breaks)
                {

                    b.Remove();

                }

                mainPart.Document.Save();

            }

        }
    }
}
