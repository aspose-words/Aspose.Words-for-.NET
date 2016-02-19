// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithBookmarks();

Document doc = new Document(dataDir + "Bookmarks.doc");

// By index.
Bookmark bookmark1 = doc.Range.Bookmarks[0];
           
// By name.
Bookmark bookmark2 = doc.Range.Bookmarks["Bookmark2"];
