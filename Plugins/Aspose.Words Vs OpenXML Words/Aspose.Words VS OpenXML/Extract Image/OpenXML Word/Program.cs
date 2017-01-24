using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.Drawing;
namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string File = FilePath + "Extract Image - OpenXML.docx";

            using (WordprocessingDocument doc = WordprocessingDocument.Open(File, false))
            {
                int imgCount = doc.MainDocumentPart.GetPartsCountOfType<ImagePart>();

                if (imgCount > 0)
                {
                    List<ImagePart> imgParts = new List<ImagePart>(doc.MainDocumentPart.ImageParts);

                    foreach (ImagePart imgPart in imgParts)
                    {
                        Image img = Image.FromStream(imgPart.GetStream());
                        string ImgfileName = imgPart.Uri.OriginalString.Substring(imgPart.Uri.OriginalString.LastIndexOf("/") + 1);

                        img.Save(FilePath + ImgfileName);
                    }
                }
            }
        }
    }
}
