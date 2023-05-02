// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document(MyDir + "Table spanning two pages.docx");

Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

// Disable breaking across pages for all rows in the table.
foreach (Row row in table.Rows)
    row.RowFormat.AllowBreakAcrossPages = false;

doc.Save(ArtifactsDir + "WorkingWithTables.RowFormatDisableBreakAcrossPages.docx");
