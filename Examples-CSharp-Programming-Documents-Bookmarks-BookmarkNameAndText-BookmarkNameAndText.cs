// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithBookmarks();

Document doc = new Document(dataDir + "Bookmark.doc");

// Use the indexer of the Bookmarks collection to obtain the desired bookmark.
Bookmark bookmark = doc.Range.Bookmarks["MyBookmark"];

// Get the name and text of the bookmark.
string name = bookmark.Name;
string text = bookmark.Text;

// Set the name and text of the bookmark.
bookmark.Name = "RenamedBookmark";
bookmark.Text = "This is a new bookmarked text.";
