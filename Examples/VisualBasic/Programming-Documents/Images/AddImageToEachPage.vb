Imports Microsoft.VisualBasic
Imports System
Imports System.IO

Imports Aspose.Words
Imports Aspose.Words.Layout
Imports System.Collections
Imports Aspose.Words.Drawing

Public Class AddImageToEachPage
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithImages()
        Dim fileName As String = "TestFile.doc"
        ' This a document that we want to add an image and custom text for each page without using the header or footer.
        Dim doc As New Document(dataDir & fileName)

        ' Create and attach collector before the document before page layout is built.
        Dim layoutCollector As New LayoutCollector(doc)

        ' Images in a document are added to paragraphs, so to add an image to every page we need to find at any paragraph 
        ' belonging to each page.
        Dim enumerator As IEnumerator = doc.SelectNodes("//Body/Paragraph").GetEnumerator()

        ' Loop through each document page.
        For page As Integer = 1 To doc.PageCount
            Do While enumerator.MoveNext()
                ' Check if the current paragraph belongs to the target page.
                Dim paragraph As Paragraph = CType(enumerator.Current, Paragraph)
                If layoutCollector.GetStartPageIndex(paragraph) = page Then
                    AddImageToPage(paragraph, page, dataDir)
                    Exit Do
                End If
            Loop
        Next page

        ' Call UpdatePageLayout() method if file is to be saved as PDF or image format
        doc.UpdatePageLayout()

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        doc.Save(dataDir)

        Console.WriteLine(vbNewLine & "Inserted images on each page of the document successfully." & vbNewLine & "File saved at " + dataDir)
    End Sub

    ''' <summary>
    ''' Adds an image to a page using the supplied paragraph.
    ''' </summary>
    ''' <param name="para">The paragraph to an an image to.</param>
    ''' <param name="page">The page number the paragraph appears on.</param>
    Public Shared Sub AddImageToPage(ByVal para As Paragraph, ByVal page As Integer, dataDir As String)
        Dim doc As Document = CType(para.Document, Document)

        Dim builder As New DocumentBuilder(doc)
        builder.MoveTo(para)

        ' Add a logo to the top left of the page. The image is placed infront of all other text.
        Dim shape As Shape = builder.InsertImage(dataDir & "Aspose Logo.png", RelativeHorizontalPosition.Page, 60, RelativeVerticalPosition.Page, 60, -1, -1, WrapType.None)

        ' Add a textbox next to the image which contains some text consisting of the page number. 
        Dim textBox As New Shape(doc, ShapeType.TextBox)

        ' We want a floating shape relative to the page.
        textBox.WrapType = WrapType.None
        textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page
        textBox.RelativeVerticalPosition = RelativeVerticalPosition.Page

        ' Set the textbox position.
        textBox.Height = 30
        textBox.Width = 200
        textBox.Left = 150
        textBox.Top = 80

        ' Add the textbox and set text.
        textBox.AppendChild(New Paragraph(doc))
        builder.InsertNode(textBox)
        builder.MoveTo(textBox.FirstChild)
        builder.Writeln("This is a custom note for page " & page)
    End Sub
End Class
