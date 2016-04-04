// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithSections();

Document srcDoc = new Document(dataDir + "Document.doc");
Document dstDoc = new Document();

Section sourceSection = srcDoc.Sections[0];
Section newSection = (Section)dstDoc.ImportNode(sourceSection, true);
dstDoc.Sections.Add(newSection);
dataDir = dataDir + "Document.Copy_out_.doc";
dstDoc.Save(dataDir);
