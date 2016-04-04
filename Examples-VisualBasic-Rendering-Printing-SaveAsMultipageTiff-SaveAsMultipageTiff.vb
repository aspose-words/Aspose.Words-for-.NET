' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()

' Open the document.
Dim doc As New Document(dataDir & "TestFile Multipage TIFF.doc")

' Save the document as multipage TIFF.
doc.Save(dataDir & "TestFile Multipage TIFF_out_.tiff")

'Create an ImageSaveOptions object to pass to the Save method
Dim options As New ImageSaveOptions(SaveFormat.Tiff)
options.PageIndex = 0
options.PageCount = 2
options.TiffCompression = TiffCompression.Ccitt4
options.Resolution = 160

doc.Save(dataDir & "TestFileWithOptions_out_.tiff", options)
