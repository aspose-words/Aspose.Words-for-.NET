// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithSections();
Document doc = new Document(dataDir + "Document.doc");
Section section = doc.Sections[0];
section.PageSetup.LeftMargin = 90; // 3.17 cm
section.PageSetup.RightMargin = 90; // 3.17 cm
section.PageSetup.TopMargin = 72; // 2.54 cm
section.PageSetup.BottomMargin = 72; // 2.54 cm
section.PageSetup.HeaderDistance = 35.4; // 1.25 cm
section.PageSetup.FooterDistance = 35.4; // 1.25 cm
section.PageSetup.TextColumns.Spacing = 35.4; // 1.25 cm
