' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_JoiningAndAppending()
Dim fileName As String = "TestFile.Destination.doc"

Dim dstDoc As New Document(dataDir & fileName)
Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")


Dim mode As ImportFormatMode = ImportFormatMode.KeepSourceFormatting

' Loop through all sections in the source document. 
' Section nodes are immediate children of the Document node so we can just enumerate the Document.
For Each srcSection As Section In srcDoc
    ' Because we are copying a section from one document to another, 
    ' it is required to import the Section node into the destination document.
    ' This adjusts any document-specific references to styles, lists, etc.
    '
    ' Importing a node creates a copy of the original node, but the copy
    ' is ready to be inserted into the destination document.
    Dim dstSection As Node = dstDoc.ImportNode(srcSection, True, mode)

    ' Now the new section node can be appended to the destination document.
    dstDoc.AppendChild(dstSection)
Next

dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
' Save the joined document
dstDoc.Save(dataDir)
