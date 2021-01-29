using System;
using Aspose.Words;
using System.IO;

namespace _01._07_CheckFormatCompatibility
{
    class Program
    {
        static void Main(string[] args)
        {
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
                        Console.WriteLine("\tPre-Microsoft Word 95 format.");
                        break;
                    case LoadFormat.Chm:
                        break;
                    case LoadFormat.FlatOpc:
                    case LoadFormat.FlatOpcMacroEnabled:
                    case LoadFormat.FlatOpcTemplate:
                    case LoadFormat.FlatOpcTemplateMacroEnabled:
                        Console.WriteLine("\tOffice Open XML WordprocessingML.");
                        break;
                    case LoadFormat.Markdown:
                        Console.WriteLine("\tMarkdown text.");
                        break;
                    case LoadFormat.Mobi:
                        Console.WriteLine("\tMOBI eBook.");
                        break;
                    case LoadFormat.Text:
                        Console.WriteLine("\tPlaintext.");
                        break;
                    case LoadFormat.Pdf:
                        Console.WriteLine("\tPDF.");
                        break;
                    default:
                        Console.WriteLine("\tUnknown format.");
                        break;
                }
            }
        }
    }
}
