// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document(MyDir + "Tables.docx");

Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
table.AllowCellSpacing = true;
table.CellSpacing = 2;
            
doc.Save(ArtifactsDir + "WorkingWithTableStylesAndFormatting.AllowCellSpacing.docx");
