// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_JoiningAndAppending();
string fileName = "TestFile.Destination.doc";
Document dstDoc = new Document(dataDir + fileName);
Document srcDoc = new Document(dataDir + "TestFile.Source.doc");
ImportFormatMode mode = ImportFormatMode.KeepSourceFormatting;

// Loop through all sections in the source document. 
// Section nodes are immediate children of the Document node so we can just enumerate the Document.
foreach (Section srcSection in srcDoc)
{
    // Because we are copying a section from one document to another, 
    // it is required to import the Section node into the destination document.
    // This adjusts any document-specific references to styles, lists, etc.
    //
    // Importing a node creates a copy of the original node, but the copy
    // is ready to be inserted into the destination document.
    Node dstSection = dstDoc.ImportNode(srcSection, true, mode);

    // Now the new section node can be appended to the destination document.
    dstDoc.AppendChild(dstSection);
}

dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
// Save the joined document
dstDoc.Save(dataDir);
