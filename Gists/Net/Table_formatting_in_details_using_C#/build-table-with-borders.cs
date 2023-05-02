// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document(MyDir + "Tables.docx");

Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            
// Clear any existing borders from the table.
table.ClearBorders();
            
// Set a green border around and inside the table.
table.SetBorders(LineStyle.Single, 1.5, Color.Green);

doc.Save(ArtifactsDir + "WorkingWithTableStylesAndFormatting.BuildTableWithBorders.docx");
