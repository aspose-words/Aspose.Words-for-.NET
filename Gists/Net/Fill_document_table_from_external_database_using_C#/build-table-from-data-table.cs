// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document();
// We can position where we want the table to be inserted and specify any extra formatting to the table.
DocumentBuilder builder = new DocumentBuilder(doc);

// We want to rotate the page landscape as we expect a wide table.
doc.FirstSection.PageSetup.Orientation = Orientation.Landscape;

DataSet ds = new DataSet();
ds.ReadXml(MyDir + "List of people.xml");
// Retrieve the data from our data source, which is stored as a DataTable.
DataTable dataTable = ds.Tables[0];

// Build a table in the document from the data contained in the DataTable.
Table table = ImportTableFromDataTable(builder, dataTable, true);

// We can apply a table style as a very quick way to apply formatting to the entire table.
table.StyleIdentifier = StyleIdentifier.MediumList2Accent1;
table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands | TableStyleOptions.LastColumn;

// For our table, we want to remove the heading for the image column.
table.FirstRow.LastCell.RemoveAllChildren();

doc.Save(ArtifactsDir + "WorkingWithTables.BuildTableFromDataTable.docx");
