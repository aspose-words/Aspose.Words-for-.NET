// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Document docA = new Document(dataDir + "TestFile.doc");
Document docB = new Document(dataDir + "TestFile - Copy.doc");
// docA now contains changes as revisions. 
docA.Compare(docB, "user", DateTime.Now);
if (docA.Revisions.Count == 0)
    Console.WriteLine("Documents are equal");
else
    Console.WriteLine("Documents are not equal");
