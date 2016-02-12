// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithSections();
Document doc = new Document(dataDir + "Section.AppendContent.doc");
// This is the section that we will append and prepend to.
Section section = doc.Sections[2];

// This copies content of the 1st section and inserts it at the beginning of the specified section.
Section sectionToPrepend = doc.Sections[0];
section.PrependContent(sectionToPrepend);

// This copies content of the 2nd section and inserts it at the end of the specified section.
Section sectionToAppend = doc.Sections[1];
section.AppendContent(sectionToAppend);
