using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;

namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string fileName = FilePath + "Remove Section Breaks.docx";
            RemoveSectionBreaks(fileName);
        }
        static void RemoveSectionBreaks(string filename)
        {

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filename, true))
            {

                MainDocumentPart mainPart = myDoc.MainDocumentPart;

                List<ParagraphProperties> paraProps = mainPart.Document.Descendants<ParagraphProperties>()

                .Where(pPr => IsSectionProps(pPr)).ToList();

                foreach (ParagraphProperties pPr in paraProps)
                {

                    pPr.RemoveChild<SectionProperties>(pPr.GetFirstChild<SectionProperties>());

                }

                mainPart.Document.Save();

            }

        }

        static bool IsSectionProps(ParagraphProperties pPr)
        {

            SectionProperties sectPr = pPr.GetFirstChild<SectionProperties>();

            if (sectPr == null)

                return false;

            else

                return true;

        }

    }
}
