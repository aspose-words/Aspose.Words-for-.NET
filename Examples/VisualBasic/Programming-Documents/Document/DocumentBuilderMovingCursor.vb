Imports System.Drawing
Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports Aspose.Words.Drawing.Charts
Imports Aspose.Words.Fields
Imports Aspose.Words.Tables
Class DocumentBuilderMovingCursor
    Public Shared Sub Run()

        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        CursorPosition(dataDir)
        MoveToNode(dataDir)
        MoveToDocumentStartEnd(dataDir)
        MoveToSection(dataDir)
        HeadersAndFooters(dataDir)
        MoveToParagraph(dataDir)
        MoveToTableCell(dataDir)
        MoveToBookmark(dataDir)
        MoveToBookmarkEnd(dataDir)
        MoveToMergeField(dataDir)

    End Sub
    Public Shared Sub CursorPosition(dataDir As String)
        ' ExStart:DocumentBuilderCursorPosition
        ' Shows how to access the current node in a document builder.
        Dim doc As Document = New Aspose.Words.Document(dataDir & Convert.ToString("DocumentBuilder.doc"))
        Dim builder As New DocumentBuilder(doc)

        Dim curNode As Node = builder.CurrentNode
        Dim curParagraph As Paragraph = builder.CurrentParagraph
        ' ExEnd:DocumentBuilderCursorPosition
        Console.WriteLine(vbLf & "Cursor move to paragraph: " + curParagraph.GetText())
    End Sub
    Public Shared Sub MoveToNode(dataDir As String)
        ' ExStart:DocumentBuilderMoveToNode
        Dim doc As New Document(dataDir & Convert.ToString("DocumentBuilder.doc"))
        Dim builder As New DocumentBuilder(doc)
        builder.MoveTo(doc.FirstSection.Body.LastParagraph)
        ' ExEnd:DocumentBuilderMoveToNode   
        Console.WriteLine(vbLf & "Cursor move to required node.")
    End Sub
    Public Shared Sub MoveToDocumentStartEnd(dataDir As String)
        ' ExStart:DocumentBuilderMoveToDocumentStartEnd
        Dim doc As New Document(dataDir & Convert.ToString("DocumentBuilder.doc"))
        Dim builder As New DocumentBuilder(doc)

        builder.MoveToDocumentEnd()
        Console.WriteLine(vbLf & "This is the end of the document.")

        builder.MoveToDocumentStart()
        Console.WriteLine(vbLf & "This is the beginning of the document.")
        ' ExEnd:DocumentBuilderMoveToDocumentStartEnd            
    End Sub
    Public Shared Sub MoveToSection(dataDir As String)
        ' ExStart:DocumentBuilderMoveToSection
        Dim doc As New Document(dataDir & Convert.ToString("DocumentBuilder.doc"))
        Dim builder As New DocumentBuilder(doc)

        ' Parameters are 0-index. Moves to third section.
        builder.MoveToSection(2)
        builder.Writeln("This is the 3rd section.")
        ' ExEnd:DocumentBuilderMoveToSection               
    End Sub
    Public Shared Sub HeadersAndFooters(dataDir As String)
        ' ExStart:DocumentBuilderHeadersAndFooters
        ' Create a blank document.
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        ' Specify that we want headers and footers different for first, even and odd pages.
        builder.PageSetup.DifferentFirstPageHeaderFooter = True
        builder.PageSetup.OddAndEvenPagesHeaderFooter = True

        ' Create the headers.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst)
        builder.Write("Header First")
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven)
        builder.Write("Header Even")
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary)
        builder.Write("Header Odd")

        ' Create three pages in the document.
        builder.MoveToSection(0)
        builder.Writeln("Page1")
        builder.InsertBreak(BreakType.PageBreak)
        builder.Writeln("Page2")
        builder.InsertBreak(BreakType.PageBreak)
        builder.Writeln("Page3")

        dataDir = dataDir & Convert.ToString("DocumentBuilder.HeadersAndFooters_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderHeadersAndFooters   
        Console.WriteLine(Convert.ToString(vbLf & "Headers and footers created successfully using DocumentBuilder." & vbLf & "File saved at ") & dataDir)
    End Sub
    Public Shared Sub MoveToParagraph(dataDir As String)
        ' ExStart:DocumentBuilderMoveToParagraph
        Dim doc As New Document(dataDir & Convert.ToString("DocumentBuilder.doc"))
        Dim builder As New DocumentBuilder(doc)

        ' Parameters are 0-index. Moves to third paragraph.
        builder.MoveToParagraph(2, 0)
        builder.Writeln("This is the 3rd paragraph.")
        ' ExEnd:DocumentBuilderMoveToParagraph               
    End Sub
    Public Shared Sub MoveToTableCell(dataDir As String)
        ' ExStart:DocumentBuilderMoveToTableCell
        Dim doc As New Document(dataDir & Convert.ToString("DocumentBuilder.doc"))
        Dim builder As New DocumentBuilder(doc)

        ' All parameters are 0-index. Moves to the 2nd table, 3rd row, 5th cell.
        builder.MoveToCell(1, 2, 4, 0)
        builder.Writeln("Hello World!")
        ' ExEnd:DocumentBuilderMoveToTableCell               
    End Sub
    Public Shared Sub MoveToBookmark(dataDir As String)
        ' ExStart:DocumentBuilderMoveToBookmark
        Dim doc As New Document(dataDir & Convert.ToString("DocumentBuilder.doc"))
        Dim builder As New DocumentBuilder(doc)

        builder.MoveToBookmark("CoolBookmark")
        builder.Writeln("This is a very cool bookmark.")
        ' ExEnd:DocumentBuilderMoveToBookmark               
    End Sub
    Public Shared Sub MoveToBookmarkEnd(dataDir As String)
        ' ExStart:DocumentBuilderMoveToBookmarkEnd
        Dim doc As New Document(dataDir & Convert.ToString("DocumentBuilder.doc"))
        Dim builder As New DocumentBuilder(doc)

        builder.MoveToBookmark("CoolBookmark", False, True)
        builder.Writeln("This is a very cool bookmark.")
        ' ExEnd:DocumentBuilderMoveToBookmarkEnd              
    End Sub
    Public Shared Sub MoveToMergeField(dataDir As String)
        ' ExStart:DocumentBuilderMoveToMergeField
        Dim doc As New Document(dataDir & Convert.ToString("DocumentBuilder.doc"))
        Dim builder As New DocumentBuilder(doc)

        builder.MoveToMergeField("NiceMergeField")
        builder.Writeln("This is a very nice merge field.")
        ' ExEnd:DocumentBuilderMoveToMergeField              
    End Sub
End Class

