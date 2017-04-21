Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports Aspose.Words.Rendering
Imports Aspose.Words
Imports System.Windows.Forms

Class PrintMultiplePagesOnOneSheet
    Public Shared Sub Run()
        ' ExStart:PrintMultiplePagesOnOneSheet
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()
        ' Open the document.
        Dim doc As New Document(dataDir & Convert.ToString("TestFile.doc"))
        ' ExStart:PrintDialogSettings
        Dim printDlg As New PrintDialog()
        ' Initialize the Print Dialog with the number of pages in the document.
        printDlg.AllowSomePages = True
        printDlg.PrinterSettings.MinimumPage = 1
        printDlg.PrinterSettings.MaximumPage = doc.PageCount
        printDlg.PrinterSettings.FromPage = 1
        printDlg.PrinterSettings.ToPage = doc.PageCount
        ' ExEnd:PrintDialogSettings
        ' Check if user accepted the print settings and proceed to preview.
        ' ExStart:CheckPrintSettings
        If Not printDlg.ShowDialog().Equals(DialogResult.OK) Then
            Return
        End If
        ' ExEnd:CheckPrintSettings
        ' Pass the printer settings from the dialog to the print document.
        Dim awPrintDoc As New MultipagePrintDocument(doc, 4, True)
        awPrintDoc.PrinterSettings = printDlg.PrinterSettings
        ' ExStart:ActivePrintPreviewDialog
        ' Create and configure the the ActivePrintPreviewDialog class.
        Dim previewDlg As New ActivePrintPreviewDialog()
        previewDlg.Document = awPrintDoc
        ' Specify additional parameters of the Print Preview dialog.
        previewDlg.ShowInTaskbar = True
        previewDlg.MinimizeBox = True
        previewDlg.Document.DocumentName = "TestFile.doc"
        previewDlg.WindowState = FormWindowState.Maximized
        ' Show appropriately configured Print Preview dialog.
        previewDlg.ShowDialog()
        ' ExEnd:ActivePrintPreviewDialog
        ' ExEnd:PrintMultiplePagesOnOneSheet
    End Sub

End Class
' ExStart:MultipagePrintDocument
Friend Class MultipagePrintDocument
    Inherits PrintDocument
    ' ExEnd:MultipagePrintDocument
    ' The data and state fields of the custom PrintDocument class.
    ' ExStart:DataAndStaticFields        
    Private ReadOnly mDocument As Document
    Private ReadOnly mPagesPerSheet As Integer
    Private ReadOnly mPrintPageBorders As Boolean
    Private mPaperSize As Size
    Private mCurrentPage As Integer
    Private mPageTo As Integer
    ' ExEnd:DataAndStaticFields
   
    ' ExStart:MultipagePrintDocumentConstructor 
    Public Sub New(document As Document, pagesPerSheet As Integer, printPageBorders As Boolean)
        If document Is Nothing Then
            Throw New ArgumentNullException("document")
        End If

        mDocument = document
        mPagesPerSheet = pagesPerSheet
        mPrintPageBorders = printPageBorders
    End Sub
    ' ExEnd:MultipagePrintDocumentConstructor    
    ' ExStart:OnBeginPrint
    Protected Overrides Sub OnBeginPrint(e As PrintEventArgs)
        MyBase.OnBeginPrint(e)

        Select Case PrinterSettings.PrintRange
            Case PrintRange.AllPages
                mCurrentPage = 0
                mPageTo = mDocument.PageCount - 1
                Exit Select
            Case PrintRange.SomePages
                mCurrentPage = PrinterSettings.FromPage - 1
                mPageTo = PrinterSettings.ToPage - 1
                Exit Select
            Case Else
                Throw New InvalidOperationException("Unsupported print range.")
        End Select

        ' Store the size of the paper, selected by user, taking into account the paper orientation.
        If PrinterSettings.DefaultPageSettings.Landscape Then
            mPaperSize = New Size(PrinterSettings.DefaultPageSettings.PaperSize.Height, PrinterSettings.DefaultPageSettings.PaperSize.Width)
        Else
            mPaperSize = New Size(PrinterSettings.DefaultPageSettings.PaperSize.Width, PrinterSettings.DefaultPageSettings.PaperSize.Height)
        End If

    End Sub
    ' ExEnd:OnBeginPrint    
    ' ExStart:OnPrintPage
    Protected Overrides Sub OnPrintPage(ByVal e As PrintPageEventArgs)
        MyBase.OnPrintPage(e)

        ' Transfer to the point units.
        e.Graphics.PageUnit = GraphicsUnit.Point
        e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit

        ' Get the number of the thumbnail placeholders across and down the paper sheet.
        Dim thumbCount As Size = GetThumbCount(mPagesPerSheet)

        ' Calculate the size of each thumbnail placeholder in points.
        ' Papersize in .NET is represented in hundreds of an inch. We need to convert this value to points first.
        Dim thumbSize As New SizeF(HundredthsInchToPoint(mPaperSize.Width) / thumbCount.Width, HundredthsInchToPoint(mPaperSize.Height) / thumbCount.Height)

        ' Select the number of the last page to be printed on this sheet of paper.
        Dim pageTo As Integer = System.Math.Min(mCurrentPage + mPagesPerSheet - 1, mPageTo)

        ' Loop through the selected pages from the stored current page to calculated last page.
        For pageIndex As Integer = mCurrentPage To pageTo
            ' Calculate the column and row indices.
            Dim columnIdx As Integer
            Dim rowIdx As Integer = System.Math.DivRem(pageIndex - mCurrentPage, thumbCount.Width, columnIdx)

            ' Define the thumbnail location in world coordinates (points in this case).
            Dim thumbLeft As Single = columnIdx * thumbSize.Width
            Dim thumbTop As Single = rowIdx * thumbSize.Height
            ' Render the document page to the Graphics object using calculated coordinates and thumbnail placeholder size.
            ' The useful return value is the scale at which the page was rendered.
            Dim scale As Single = mDocument.RenderToSize(pageIndex, e.Graphics, thumbLeft, thumbTop, thumbSize.Width, thumbSize.Height)

            ' Draw the page borders (the page thumbnail could be smaller than the thumbnail placeholder size).
            If mPrintPageBorders Then
                ' Get the real 100% size of the page in points.
                Dim pageSize As SizeF = mDocument.GetPageInfo(pageIndex).SizeInPoints
                ' Draw the border around the scaled page using the known scale factor.
                e.Graphics.DrawRectangle(Pens.Black, thumbLeft, thumbTop, pageSize.Width * scale, pageSize.Height * scale)

                ' Draws the border around the thumbnail placeholder.
                e.Graphics.DrawRectangle(Pens.Red, thumbLeft, thumbTop, thumbSize.Width, thumbSize.Height)
            End If
        Next pageIndex

        ' Re-calculate next current page and continue with printing if such page resides within the print range.
        mCurrentPage = mCurrentPage + mPagesPerSheet
        e.HasMorePages = (mCurrentPage <= mPageTo)
    End Sub
    ' ExEnd:OnPrintPage
    ' ExStart:HundredthsInchToPoint
    Private Shared Function HundredthsInchToPoint(value As Single) As Single
        Return CSng(ConvertUtil.InchToPoint(value / 100))
    End Function
    ' ExEnd:HundredthsInchToPoint
    ' ExStart:GetThumbCount
    Private Function GetThumbCount(pagesPerSheet As Integer) As Size
        Dim size As Size
        ' Define the number of the columns and rows on the sheet for the Landscape-oriented paper.
        Select Case pagesPerSheet
            Case 16
                size = New Size(4, 4)
                Exit Select
            Case 9
                size = New Size(3, 3)
                Exit Select
            Case 8
                size = New Size(4, 2)
                Exit Select
            Case 6
                size = New Size(3, 2)
                Exit Select
            Case 4
                size = New Size(2, 2)
                Exit Select
            Case 2
                size = New Size(2, 1)
                Exit Select
            Case Else
                size = New Size(1, 1)
                Exit Select
        End Select
        ' Switch the width and height if the paper is in the Portrait orientation.
        If mPaperSize.Width < mPaperSize.Height Then
            Return New Size(size.Height, size.Width)
        End If
        Return size
    End Function
    ' ExEnd:GetThumbCount
End Class
