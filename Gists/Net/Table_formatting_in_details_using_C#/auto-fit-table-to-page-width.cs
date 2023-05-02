// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document(MyDir + "Tables.docx");

Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
// Autofit the first table to the page width.
table.AutoFit(AutoFitBehavior.AutoFitToWindow);

doc.Save(ArtifactsDir + "WorkingWithTables.AutoFitTableToWindow.docx");
