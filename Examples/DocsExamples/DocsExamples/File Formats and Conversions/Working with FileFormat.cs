using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions
{
    public class WorkingWithFileFormat : DocsExamplesBase
    {
        [Test]
        public void DetectFileFormat()
        {
            //ExStart:CheckFormatCompatibility
            string supportedDir = ArtifactsDir + "Supported";
            string unknownDir = ArtifactsDir + "Unknown";
            string encryptedDir = ArtifactsDir + "Encrypted";
            string pre97Dir = ArtifactsDir + "Pre97";

            // Create the directories if they do not already exist.
            if (Directory.Exists(supportedDir) == false)
                Directory.CreateDirectory(supportedDir);
            if (Directory.Exists(unknownDir) == false)
                Directory.CreateDirectory(unknownDir);
            if (Directory.Exists(encryptedDir) == false)
                Directory.CreateDirectory(encryptedDir);
            if (Directory.Exists(pre97Dir) == false)
                Directory.CreateDirectory(pre97Dir);

            //ExStart:GetListOfFilesInFolder
            IEnumerable<string> fileList = Directory.GetFiles(MyDir).Where(name => !name.EndsWith("Corrupted document.docx"));
            //ExEnd:GetListOfFilesInFolder
            foreach (string fileName in fileList)
            {
                string nameOnly = Path.GetFileName(fileName);
                
                Console.Write(nameOnly);
                //ExStart:DetectFileFormat
                FileFormatInfo info = FileFormatUtil.DetectFileFormat(fileName);

                // Display the document type
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
                        Console.WriteLine("\tUnknown format.");
                        break;
                }
                //ExEnd:DetectFileFormat

                if (info.IsEncrypted)
                {
                    Console.WriteLine("\tAn encrypted document.");
                    File.Copy(fileName, Path.Combine(encryptedDir, nameOnly), true);
                }
                else
                {
                    switch (info.LoadFormat)
                    {
                        case LoadFormat.DocPreWord60:
                            File.Copy(fileName, Path.Combine(pre97Dir, nameOnly), true);
                            break;
                        case LoadFormat.Unknown:
                            File.Copy(fileName, Path.Combine(unknownDir, nameOnly), true);
                            break;
                        default:
                            File.Copy(fileName, Path.Combine(supportedDir, nameOnly), true);
                            break;
                    }
                }
            }
            //ExEnd:CheckFormatCompatibility
        }

        [Test]
        public void DetectDocumentSignatures()
        {
            //ExStart:DetectDocumentSignatures
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(MyDir + "Digitally signed.docx");

            if (info.HasDigitalSignature)
            {
                Console.WriteLine(
                    $"Document {Path.GetFileName(MyDir + "Digitally signed.docx")} has digital signatures, " +
                    "they will be lost if you open/save this document with Aspose.Words.");
            }
            //ExEnd:DetectDocumentSignatures            
        }

        [Test]
        public void VerifyEncryptedDocument()
        {
            //ExStart:VerifyEncryptedDocument
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(MyDir + "Encrypted.docx");
            Console.WriteLine(info.IsEncrypted);
            //ExEnd:VerifyEncryptedDocument
        }
    }
}