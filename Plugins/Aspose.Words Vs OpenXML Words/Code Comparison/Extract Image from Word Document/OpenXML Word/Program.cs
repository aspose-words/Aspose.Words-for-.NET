using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
namespace OpenXML_Word
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = "Test.docx";
            using (WordprocessingDocument doc = WordprocessingDocument.Open(fileName, false))
            {
                int imgCount = doc.MainDocumentPart.GetPartsCountOfType<ImagePart>();

                if (imgCount > 0)
                {
                    List<ImagePart> imgParts = new List<ImagePart>(doc.MainDocumentPart.ImageParts);

                    foreach (ImagePart imgPart in imgParts)
                    {
                        Image img = Image.FromStream(imgPart.GetStream());
                        string ImgfileName = imgPart.Uri.OriginalString.Substring(imgPart.Uri.OriginalString.LastIndexOf("/") + 1);
                       
                        img.Save(ImgfileName);
                    }
                }
            }
        }
    }
}
