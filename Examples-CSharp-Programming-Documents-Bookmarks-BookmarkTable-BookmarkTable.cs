// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithBookmarks();
            
// Create empty document
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);

Table table = builder.StartTable();

// Insert a cell
builder.InsertCell();

// Start bookmark here after calling InsertCell
builder.StartBookmark("MyBookmark");

builder.Write("This is row 1 cell 1");

// Insert a cell
builder.InsertCell();
builder.Write("This is row 1 cell 2");

builder.EndRow();

// Insert a cell
builder.InsertCell();
builder.Writeln("This is row 2 cell 1");

// Insert a cell
builder.InsertCell();
builder.Writeln("This is row 2 cell 2");

builder.EndRow();

builder.EndTable();
// End of bookmark
builder.EndBookmark("MyBookmark");

dataDir = dataDir + "Bookmark.Table_out_.doc";
doc.Save(dataDir);
