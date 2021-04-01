using System;
using System.IO;
using Aspose.Words.Drawing;

namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
            // Check for license and apply if exists
            string licenseFile = AppDomain.CurrentDomain.BaseDirectory + "Aspose.Words.lic";
            if (File.Exists(licenseFile))
            {
                // Apply Aspose.Words API License
                Aspose.Words.License license = new Aspose.Words.License();
                // Place license file in Bin/Debug/ Folder
                license.SetLicense("Aspose.Words.lic");
            }

            string filePath = Path.GetDirectoryName(Path.GetDirectoryName(Directory.GetCurrentDirectory())) + @"\data\" + "Extract Images from Word Document.doc";

            Document wordDocument = new Document(filePath);
            
            NodeCollection pictures = wordDocument.GetChildNodes(NodeType.Shape, true);
            int imageindex = 0;
            foreach (Shape shape in pictures)
            {
                string imageType = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                if (shape.HasImage)
                {
                    string imageFileName = "Aspose_" + (imageindex++).ToString() + "_" + shape.Name + imageType;
                    shape.ImageData.Save(imageFileName);
                }
            }

        }
    }
}
