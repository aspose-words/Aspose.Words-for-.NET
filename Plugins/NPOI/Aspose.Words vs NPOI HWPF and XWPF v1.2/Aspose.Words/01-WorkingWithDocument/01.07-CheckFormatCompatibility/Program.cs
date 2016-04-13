using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using System.IO;

namespace _01._07_CheckFormatCompatibility
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


            string dataPath = "../../data/";
            string[] fileList = Directory.GetFiles(dataPath);

            // Loop through all found files.
            foreach (string filePath in fileList)
            {
                FileInfo file = new FileInfo(filePath);

                // Extract and display the file name without the path.
                String nameOnly = file.Name;
                Console.WriteLine(nameOnly);

                // Check the file format and move the file to the appropriate folder.
                String fileName = file.FullName;
                FileFormatInfo info = FileFormatUtil.DetectFileFormat(fileName);

                // Display the document type.
                switch (info.LoadFormat)
                {
                    case LoadFormat.Doc:
                        Console.WriteLine("\tMicrosoft Word 97-2003 document.");
                        break;
                    case LoadFormat.Dot:
                        Console.WriteLine("\tMicrosoft Word 97-2003 template.");
                        break;
                    case LoadFormat.Docx:
                        Console.WriteLine("\tOffice Open XML WordprocessingML Macro-Free Document.");
                        break;
                    case LoadFormat.Docm:
                        Console.WriteLine("\tOffice Open XML WordprocessingML Macro-Enabled Document.");
                        break;
                    case LoadFormat.Dotx:
                        Console.WriteLine("\tOffice Open XML WordprocessingML Macro-Free Template.");
                        break;
                    case LoadFormat.Dotm:
                        Console.WriteLine("\tOffice Open XML WordprocessingML Macro-Enabled Template.");
                        break;
                    case LoadFormat.FlatOpc:
                        Console.WriteLine("\tFlat OPC document.");
                        break;
                    case LoadFormat.Rtf:
                        Console.WriteLine("\tRTF format.");
                        break;
                    case LoadFormat.WordML:
                        Console.WriteLine("\tMicrosoft Word 2003 WordprocessingML format.");
                        break;
                    case LoadFormat.Html:
                        Console.WriteLine("\tHTML format.");
                        break;
                    case LoadFormat.Mhtml:
                        Console.WriteLine("\tMHTML (Web archive) format.");
                        break;
                    case LoadFormat.Odt:
                        Console.WriteLine("\tOpenDocument Text.");
                        break;
                    case LoadFormat.Ott:
                        Console.WriteLine("\tOpenDocument Text Template.");
                        break;
                    case LoadFormat.DocPreWord60:
                        Console.WriteLine("\tMS Word 6 or Word 95 format.");
                        break;
                    case LoadFormat.Unknown:
                    default:
                        Console.WriteLine("\tUnknown format.");
                        break;
                }
            }
        }
    }
}
