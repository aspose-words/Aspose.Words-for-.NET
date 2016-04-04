' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim docA As New Document(dataDir & Convert.ToString("TestFile.doc"))
Dim docB As New Document(dataDir & Convert.ToString("TestFile - Copy.doc"))
' docA now contains changes as revisions. 
docA.Compare(docB, "user", DateTime.Now)
