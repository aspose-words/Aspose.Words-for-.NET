// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document(MyDir + "Tables.docx");

Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
// Disable autofitting on this table.
table.AutoFit(AutoFitBehavior.FixedColumnWidths);

doc.Save(ArtifactsDir + "WorkingWithTables.AutoFitTableToFixedColumnWidths.docx");
