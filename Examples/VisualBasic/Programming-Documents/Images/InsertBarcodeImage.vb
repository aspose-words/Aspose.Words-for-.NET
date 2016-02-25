Imports System
Imports System.Collections
Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Tables
Imports Aspose.Words.Fields
Imports Aspose.Words.Layout
Imports Aspose.Words.Drawing

Public Class InsertBarcodeImage
    Public Shared Sub Run()
        ' ExStart:InsertBarcodeImage
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithImages()
        ' Create a blank documenet.
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        ' The number of pages the document should have.
        Dim numPages As Integer = 4
        ' The document starts with one section, insert the barcode into this existing section.
        InsertBarcodeIntoFooter(builder, doc.FirstSection, 1, HeaderFooterType.FooterPrimary)

        For i As Integer = 1 To numPages - 1
            ' Clone the first section and add it into the end of the document.
            Dim cloneSection As Section = DirectCast(doc.FirstSection.Clone(False), Section)
            cloneSection.PageSetup.SectionStart = SectionStart.NewPage
            doc.AppendChild(cloneSection)

            ' Insert the barcode and other information into the footer of the section.
            InsertBarcodeIntoFooter(builder, cloneSection, i, HeaderFooterType.FooterPrimary)
        Next

        dataDir = dataDir & Convert.ToString("Document_out_.docx")
        ' Save the document as a PDF to disk. You can also save this directly to a stream.
        doc.Save(dataDir)
        ' ExEnd:InsertBarcodeImage
        Console.WriteLine(Convert.ToString(vbLf & "Barcode image on each page of document inserted successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
    ' ExStart:InsertBarcodeIntoFooter
    Private Shared Sub InsertBarcodeIntoFooter(builder As DocumentBuilder, section As Section, pageId As Integer, footerType As HeaderFooterType)
        ' Move to the footer type in the specific section.
        builder.MoveToSection(section.Document.IndexOf(section))
        builder.MoveToHeaderFooter(footerType)

        ' Insert the barcode, then move to the next line and insert the ID along with the page number.
        ' Use pageId if you need to insert a different barcode on each page. 0 = First page, 1 = Second page etc.    
        builder.InsertImage(System.Drawing.Image.FromFile(RunExamples.GetDataDir_WorkingWithImages() + "Barcode1.png"))
        builder.Writeln()
        builder.Write("1234567890")
        builder.InsertField("PAGE")

        ' Create a right aligned tab at the right margin.
        Dim tabPos As Double = section.PageSetup.PageWidth - section.PageSetup.RightMargin - section.PageSetup.LeftMargin
        builder.CurrentParagraph.ParagraphFormat.TabStops.Add(New TabStop(tabPos, TabAlignment.Right, TabLeader.None))

        ' Move to the right hand side of the page and insert the page and page total.
        builder.Write(ControlChar.Tab)
        builder.InsertField("PAGE")
        builder.Write(" of ")
        builder.InsertField("NUMPAGES")
    End Sub
    ' ExEnd:InsertBarcodeIntoFooter
End Class
