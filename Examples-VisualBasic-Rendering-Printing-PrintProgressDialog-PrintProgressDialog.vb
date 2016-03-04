' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()
' Load the documents which store the shapes we want to render.
Dim doc As New Document(dataDir & Convert.ToString("TestFile RenderShape.doc"))
' Obtain the settings of the default printer
Dim settings As New System.Drawing.Printing.PrinterSettings()

' The standard print controller comes with no UI
Dim standardPrintController As System.Drawing.Printing.PrintController = New System.Drawing.Printing.StandardPrintController()

' Print the document using the custom print controller
Dim prntDoc As New AsposeWordsPrintDocument(doc)
prntDoc.PrinterSettings = settings
prntDoc.PrintController = standardPrintController
prntDoc.Print()
