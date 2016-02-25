' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
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
