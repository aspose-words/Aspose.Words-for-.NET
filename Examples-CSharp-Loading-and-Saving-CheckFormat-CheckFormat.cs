// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

string supportedDir = dataDir + "OutSupported";
string unknownDir = dataDir + "OutUnknown";
string encryptedDir = dataDir + "OutEncrypted";
string pre97Dir = dataDir + "OutPre97";

// Create the directories if they do not already exist
if (Directory.Exists(supportedDir) == false)
    Directory.CreateDirectory(supportedDir);
if (Directory.Exists(unknownDir) == false)
    Directory.CreateDirectory(unknownDir);
if (Directory.Exists(encryptedDir) == false)
    Directory.CreateDirectory(encryptedDir);
if (Directory.Exists(pre97Dir) == false)
    Directory.CreateDirectory(pre97Dir);

string[] fileList = Directory.GetFiles(dataDir);
// Loop through all found files.
foreach (string fileName in fileList)
{
    // Extract and display the file name without the path.
    string nameOnly = Path.GetFileName(fileName);
    Console.Write(nameOnly);

    // Check the file format and move the file to the appropriate folder.
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
        case LoadFormat.DocPreWord97:
            Console.WriteLine("\tMS Word 6 or Word 95 format.");
            break;
        case LoadFormat.Unknown:
        default:
            Console.WriteLine("\tUnknown format.");
            break;
    }

    // Now copy the document into the appropriate folder.
    if (info.IsEncrypted)
    {
        Console.WriteLine("\tAn encrypted document.");
        File.Copy(fileName, Path.Combine(encryptedDir, nameOnly), true);
    }
    else
    {
        switch (info.LoadFormat)
        {
            case LoadFormat.DocPreWord97:
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
