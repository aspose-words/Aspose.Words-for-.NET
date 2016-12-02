Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Drawing

Imports Aspose.Words
Imports Aspose.Words.Layout
Imports Aspose.Words.Rendering

Public Class EnumerateLayoutElements
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()

        Dim doc As New Document(dataDir & "TestFile.docx")

        ' This creates an enumerator which is used to "walk" the elements of a rendered document.
        Dim it As New LayoutEnumerator(doc)

        ' This sample uses the enumerator to write information about each layout element to the console.
        LayoutInfoWriter.Run(it)

        ' This sample adds a border around each layout element and saves each page as a JPEG image to the data directory.
        OutlineLayoutEntitiesRenderer.Run(doc, it, dataDir)

        Console.WriteLine(vbNewLine + "Enumerate layout elements example ran successfully.")
    End Sub
End Class

Friend Class LayoutInfoWriter
    Public Shared Sub Run(ByVal it As LayoutEnumerator)
        DisplayLayoutElements(it, String.Empty)
    End Sub

    ''' <summary>
    ''' Enumerates forward through each layout element in the document and prints out details of each element. 
    ''' </summary>
    Private Shared Sub DisplayLayoutElements(ByVal it As LayoutEnumerator, ByVal padding As String)
        Do
            DisplayEntityInfo(it, padding)

            If it.MoveFirstChild() Then
                ' Recurse into this child element.
                DisplayLayoutElements(it, AddPadding(padding))
                it.MoveParent()
            End If
        Loop While it.MoveNext()
    End Sub

    ''' <summary>
    ''' Displays information about the current layout entity to the console.
    ''' </summary>
    Private Shared Sub DisplayEntityInfo(ByVal it As LayoutEnumerator, ByVal padding As String)
        Console.Write(padding & it.Type & " - " & it.Kind)

        If it.Type = LayoutEntityType.Span Then
            Console.Write(" - " & it.Text)
        End If

        Console.WriteLine()
    End Sub

    ''' <summary>
    ''' Returns a string of spaces for padding purposes.
    ''' </summary>
    Private Shared Function AddPadding(ByVal padding As String) As String
        Return padding & New String(" "c, 4)
    End Function
End Class

Friend Class OutlineLayoutEntitiesRenderer
    Public Shared Sub Run(ByVal doc As Document, ByVal it As LayoutEnumerator, ByVal folderPath As String)
        ' Make sure the enumerator is at the beginning of the document.
        it.Reset()

        For pageIndex As Integer = 0 To doc.PageCount - 1
            ' Use the document class to find information about the current page.
            Dim pageInfo As PageInfo = doc.GetPageInfo(pageIndex)

            Const resolution As Single = 150.0F
            Dim pageSize As Size = pageInfo.GetSizeInPixels(1.0F, resolution)

            Using img As New Bitmap(pageSize.Width, pageSize.Height)
                img.SetResolution(resolution, resolution)

                Using g As Graphics = Graphics.FromImage(img)
                    ' Make the background white.
                    g.Clear(Color.White)

                    ' Render the page to the graphics.
                    doc.RenderToScale(pageIndex, g, 0.0F, 0.0F, 1.0F)

                    ' Add an outline around each element on the page using the graphics object.
                    AddBoundingBoxToElementsOnPage(it, g)

                    ' Move the enumerator to the next page if there is one.
                    it.MoveNext()

                    img.Save(folderPath & String.Format("TestFile Page {0} Out.png", pageIndex + 1))
                End Using
            End Using
        Next pageIndex
    End Sub

    ''' <summary>
    ''' Adds a colored border around each layout element on the page.
    ''' </summary>
    Private Shared Sub AddBoundingBoxToElementsOnPage(ByVal it As LayoutEnumerator, ByVal g As Graphics)
        Do
            ' This time instead of MoveFirstChild and MoveNext, we use MoveLastChild and MovePrevious to enumerate from last to first.
            ' Enumeration is done backward so the lines of child entities are drawn first and don' T overlap the lines of the parent.
            If it.MoveLastChild() Then
                AddBoundingBoxToElementsOnPage(it, g)
                it.MoveParent()
            End If

            ' Convert the rectangle representing the position of the layout entity on the page from points to pixels.
            Dim rectF As RectangleF = it.Rectangle
            Dim rect As New Rectangle(PointToPixel(rectF.Left, g.DpiX), PointToPixel(rectF.Top, g.DpiY), PointToPixel(rectF.Width, g.DpiX), PointToPixel(rectF.Height, g.DpiY))

            ' Draw a line around the layout entity on the page.
            g.DrawRectangle(GetColoredPenFromType(it.Type), rect)

            ' Stop after all elements on the page have been procesed.
            If it.Type = LayoutEntityType.Page Then
                Return
            End If

        Loop While it.MovePrevious()
    End Sub

    ''' <summary>
    ''' Returns a different colored pen for each entity type.
    ''' </summary>
    Private Shared Function GetColoredPenFromType(ByVal type As LayoutEntityType) As Pen
        Select Case type
            Case LayoutEntityType.Cell
                Return Pens.Purple
            Case LayoutEntityType.Column
                Return Pens.Green
            Case LayoutEntityType.Comment
                Return Pens.LightBlue
            Case LayoutEntityType.Endnote
                Return Pens.DarkRed
            Case LayoutEntityType.Footnote
                Return Pens.DarkBlue
            Case LayoutEntityType.HeaderFooter
                Return Pens.DarkGreen
            Case LayoutEntityType.Line
                Return Pens.Blue
            Case LayoutEntityType.NoteSeparator
                Return Pens.LightGreen
            Case LayoutEntityType.Page
                Return Pens.Red
            Case LayoutEntityType.Row
                Return Pens.Orange
            Case LayoutEntityType.Span
                Return Pens.Red
            Case LayoutEntityType.TextBox
                Return Pens.Yellow
            Case Else
                Return Pens.Red
        End Select
    End Function

    ''' <summary>
    ''' Converts a value in points to pixels.
    ''' </summary>
    Private Shared Function PointToPixel(ByVal value As Single, ByVal resolution As Double) As Integer
        Return Convert.ToInt32(ConvertUtil.PointToPixel(value, resolution))
    End Function
End Class
