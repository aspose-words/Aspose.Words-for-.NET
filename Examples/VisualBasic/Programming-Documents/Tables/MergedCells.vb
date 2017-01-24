Imports System.Collections.Generic
Imports System.IO
Imports System.Xml
Imports System.Drawing
Imports Aspose.Words
Imports Aspose.Words.Tables
Imports System.Diagnostics
Imports Aspose.Words.Saving
Public Class MergedCells
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithTables()
        CheckCellsMerged(dataDir)
        ' The below method shows how to create a table with two rows with cells in the first row horizontally merged.
        HorizontalMerge(dataDir)
        ' The below method shows how to create a table with two columns with cells merged vertically in the first column.
        VerticalMerge(dataDir)
        ' The below method shows how to merges the range of cells between the two specified cells.   
        MergeCellRange(dataDir)
        ' Show how to prints the horizontal and vertical merge of a cell.
        PrintHorizontalAndVerticalMerged(dataDir)
    End Sub
    Public Shared Sub CheckCellsMerged(dataDir As String)
        ' ExStart:CheckCellsMerged 
        Dim doc As New Document(dataDir & Convert.ToString("Table.MergedCells.doc"))

        ' Retrieve the first table in the document.
        Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)

        For Each row As Row In table.Rows
            For Each cell As Cell In row.Cells
                Console.WriteLine(PrintCellMergeType(cell))
            Next
        Next
        ' ExEnd:CheckCellsMerged 
    End Sub
    ' ExStart:PrintCellMergeType 
    Public Shared Function PrintCellMergeType(cell As Cell) As String
        Dim isHorizontallyMerged As Boolean = cell.CellFormat.HorizontalMerge <> CellMerge.None
        Dim isVerticallyMerged As Boolean = cell.CellFormat.VerticalMerge <> CellMerge.None
        Dim cellLocation As String = String.Format("R{0}, C{1}", cell.ParentRow.ParentTable.IndexOf(cell.ParentRow) + 1, cell.ParentRow.IndexOf(cell) + 1)

        If isHorizontallyMerged AndAlso isVerticallyMerged Then
            Return String.Format("The cell at {0} is both horizontally and vertically merged", cellLocation)
        ElseIf isHorizontallyMerged Then
            Return String.Format("The cell at {0} is horizontally merged.", cellLocation)
        ElseIf isVerticallyMerged Then
            Return String.Format("The cell at {0} is vertically merged", cellLocation)
        Else
            Return String.Format("The cell at {0} is not merged", cellLocation)
        End If
    End Function
    ' ExEnd:PrintCellMergeType
    Public Shared Sub VerticalMerge(dataDir As String)
        ' ExStart:VerticalMerge           
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        builder.InsertCell()
        builder.CellFormat.VerticalMerge = CellMerge.First
        builder.Write("Text in merged cells.")

        builder.InsertCell()
        builder.CellFormat.VerticalMerge = CellMerge.None
        builder.Write("Text in one cell")
        builder.EndRow()

        builder.InsertCell()
        ' This cell is vertically merged to the cell above and should be empty.
        builder.CellFormat.VerticalMerge = CellMerge.Previous

        builder.InsertCell()
        builder.CellFormat.VerticalMerge = CellMerge.None
        builder.Write("Text in another cell")
        builder.EndRow()
        builder.EndTable()
        dataDir = dataDir & Convert.ToString("Table.VerticalMerge_out.doc")

        ' Save the document to disk.
        doc.Save(dataDir)
        ' ExEnd:VerticalMerge
        Console.WriteLine(Convert.ToString(vbLf & "Table created successfully with two columns with cells merged vertically in the first column." & vbLf & "File saved at ") & dataDir)
    End Sub
    Public Shared Sub HorizontalMerge(dataDir As String)
        ' ExStart:HorizontalMerge         
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        builder.InsertCell()
        builder.CellFormat.HorizontalMerge = CellMerge.First
        builder.Write("Text in merged cells.")

        builder.InsertCell()
        ' This cell is merged to the previous and should be empty.
        builder.CellFormat.HorizontalMerge = CellMerge.Previous
        builder.EndRow()

        builder.InsertCell()
        builder.CellFormat.HorizontalMerge = CellMerge.None
        builder.Write("Text in one cell.")

        builder.InsertCell()
        builder.Write("Text in another cell.")
        builder.EndRow()
        builder.EndTable()
        dataDir = dataDir & Convert.ToString("Table.HorizontalMerge_out.doc")

        ' Save the document to disk.
        doc.Save(dataDir)
        ' ExEnd:HorizontalMerge
        Console.WriteLine(Convert.ToString(vbLf & "Table created successfully with cells in the first row horizontally merged." & vbLf & "File saved at ") & dataDir)

    End Sub
    Public Shared Sub MergeCellRange(dataDir As String)
        ' ExStart:MergeCellRange
        ' Open the document
        Dim doc As New Document(dataDir & Convert.ToString("Table.Document.doc"))

        ' Retrieve the first table in the body of the first section.
        Dim table As Table = doc.FirstSection.Body.Tables(0)

        ' We want to merge the range of cells found inbetween these two cells.
        Dim cellStartRange As Cell = table.Rows(2).Cells(2)
        Dim cellEndRange As Cell = table.Rows(3).Cells(3)

        ' Merge all the cells between the two specified cells into one.
        MergeCells(cellStartRange, cellEndRange)
        dataDir = dataDir & Convert.ToString("Table.MergeCellRange_out.doc")
        ' Save the document.
        doc.Save(dataDir)
        ' ExEnd:MergeCellRange
        Console.WriteLine(Convert.ToString(vbLf & "Cells merged successfully." & vbLf & "File saved at ") & dataDir)

    End Sub

    Public Shared Sub PrintHorizontalAndVerticalMerged(dataDir As String)
        ' ExStart:PrintHorizontalAndVerticalMerged
        Dim doc As New Document(dataDir & Convert.ToString("Table.MergedCells.doc"))

        ' Create visitor
        Dim visitor As New SpanVisitor(doc)

        ' Accept visitor
        doc.Accept(visitor)
        ' ExEnd:PrintHorizontalAndVerticalMerged
        Console.WriteLine(vbLf & "Horizontal and vertical merged of a cell prints successfully.")

    End Sub
    ' ExStart:MergeCells
    Public Shared Sub MergeCells(ByVal startCell As Cell, ByVal endCell As Cell)
        Dim parentTable As Table = startCell.ParentRow.ParentTable

        ' Find the row and cell indices for the start and end cell.
        Dim startCellPos As New Point(startCell.ParentRow.IndexOf(startCell), parentTable.IndexOf(startCell.ParentRow))
        Dim endCellPos As New Point(endCell.ParentRow.IndexOf(endCell), parentTable.IndexOf(endCell.ParentRow))
        ' Create the range of cells to be merged based off these indices. Inverse each index if the end cell if before the start cell. 
        Dim mergeRange As New Rectangle(System.Math.Min(startCellPos.X, endCellPos.X), System.Math.Min(startCellPos.Y, endCellPos.Y), System.Math.Abs(endCellPos.X - startCellPos.X) + 1, System.Math.Abs(endCellPos.Y - startCellPos.Y) + 1)

        For Each row As Row In parentTable.Rows
            For Each cell As Cell In row.Cells
                Dim currentPos As New Point(row.IndexOf(cell), parentTable.IndexOf(row))

                ' Check if the current cell is inside our merge range then merge it.
                If mergeRange.Contains(currentPos) Then
                    If currentPos.X = mergeRange.X Then
                        cell.CellFormat.HorizontalMerge = CellMerge.First
                    Else
                        cell.CellFormat.HorizontalMerge = CellMerge.Previous
                    End If

                    If currentPos.Y = mergeRange.Y Then
                        cell.CellFormat.VerticalMerge = CellMerge.First
                    Else
                        cell.CellFormat.VerticalMerge = CellMerge.Previous
                    End If
                End If
            Next cell
        Next row
    End Sub
    ' ExEnd:MergeCells
    ' ExStart:HorizontalAndVerticalMergeHelperClasses
    ''' <summary>
    ''' Helper class that contains collection of rowinfo for each row
    ''' </summary>
    Public Class TableInfo
        Public ReadOnly Property Rows() As List(Of RowInfo)
            Get
                Return mRows
            End Get
        End Property

        Private mRows As New List(Of RowInfo)()
    End Class

    ''' <summary>
    ''' Helper class that contains collection of cellinfo for each cell
    ''' </summary>
    Public Class RowInfo
        Public ReadOnly Property Cells() As List(Of CellInfo)
            Get
                Return mCells
            End Get
        End Property

        Private mCells As New List(Of CellInfo)()
    End Class

    ''' <summary>
    ''' Helper class that contains info about cell. currently here is only colspan and rowspan
    ''' </summary>
    Public Class CellInfo
        Public Sub New(colSpan As Integer, rowSpan As Integer)
            mColSpan = colSpan
            mRowSpan = rowSpan
        End Sub

        Public ReadOnly Property ColSpan() As Integer
            Get
                Return mColSpan
            End Get
        End Property

        Public ReadOnly Property RowSpan() As Integer
            Get
                Return mRowSpan
            End Get
        End Property

        Private mColSpan As Integer = 0
        Private mRowSpan As Integer = 0
    End Class

    Public Class SpanVisitor
        Inherits DocumentVisitor

        ''' <summary>
        ''' Creates new SpanVisitor instance
        ''' </summary>
        ''' <param name="doc">Is document which we should parse</param>
        Public Sub New(doc As Document)
            ' Get collection of tables from the document
            mWordTables = doc.GetChildNodes(NodeType.Table, True)

            ' Convert document to HTML
            ' We will parse HTML to determine rowspan and colspan of each cell
            Dim htmlStream As New MemoryStream()

            Dim options As New HtmlSaveOptions()
            options.ImagesFolder = Path.GetTempPath()

            doc.Save(htmlStream, options)

            ' Load HTML into the XML document
            Dim xmlDoc As New XmlDocument()
            htmlStream.Position = 0
            xmlDoc.Load(htmlStream)

            ' Get collection of tables in the HTML document
            Dim tables As XmlNodeList = xmlDoc.DocumentElement.SelectNodes("//table")

            For Each table As XmlNode In tables
                Dim tableInf As New TableInfo()
                ' Get collection of rows in the table
                Dim rows As XmlNodeList = table.SelectNodes("tr")

                For Each row As XmlNode In rows
                    Dim rowInf As New RowInfo()

                    ' Get collection of cells
                    Dim cells As XmlNodeList = row.SelectNodes("td")

                    For Each cell As XmlNode In cells
                        ' Determine row span and colspan of the current cell
                        Dim colSpanAttr As XmlAttribute = cell.Attributes("colspan")
                        Dim rowSpanAttr As XmlAttribute = cell.Attributes("rowspan")

                        Dim colSpan As Integer = If(colSpanAttr Is Nothing, 0, Int32.Parse(colSpanAttr.Value))
                        Dim rowSpan As Integer = If(rowSpanAttr Is Nothing, 0, Int32.Parse(rowSpanAttr.Value))

                        Dim cellInf As New CellInfo(colSpan, rowSpan)
                        rowInf.Cells.Add(cellInf)
                    Next

                    tableInf.Rows.Add(rowInf)
                Next

                mTables.Add(tableInf)
            Next
        End Sub

        Public Overrides Function VisitCellStart(cell As Aspose.Words.Tables.Cell) As VisitorAction
            ' Determone index of current table
            Dim tabIdx As Integer = mWordTables.IndexOf(cell.ParentRow.ParentTable)

            ' Determine index of current row
            Dim rowIdx As Integer = cell.ParentRow.ParentTable.IndexOf(cell.ParentRow)

            ' And determine index of current cell
            Dim cellIdx As Integer = cell.ParentRow.IndexOf(cell)

            ' Determine colspan and rowspan of current cell
            Dim colSpan As Integer = 0
            Dim rowSpan As Integer = 0
            If tabIdx < mTables.Count AndAlso rowIdx < mTables(tabIdx).Rows.Count AndAlso cellIdx < mTables(tabIdx).Rows(rowIdx).Cells.Count Then
                colSpan = mTables(tabIdx).Rows(rowIdx).Cells(cellIdx).ColSpan
                rowSpan = mTables(tabIdx).Rows(rowIdx).Cells(cellIdx).RowSpan
            End If

            Console.WriteLine("{0}.{1}.{2} colspan={3}" & vbTab & " rowspan={4}", tabIdx, rowIdx, cellIdx, colSpan, rowSpan)

            Return VisitorAction.[Continue]
        End Function


        Private mTables As New List(Of TableInfo)()
        Private mWordTables As NodeCollection = Nothing
    End Class
    ' ExEnd:HorizontalAndVerticalMergeHelperClasses
End Class
