Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Drawing
Imports Aspose.Words
Imports Aspose.Words.Layout
Imports Aspose.Words.Rendering
Public Class PrintProgressDialog
    Public Shared Sub Run()
        ' ExStart:PrintProgressDialog
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
        ' ExEnd:PrintProgressDialog
    End Sub
End Class
