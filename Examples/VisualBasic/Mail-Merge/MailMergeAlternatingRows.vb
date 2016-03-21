Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.IO
Imports Aspose.Words
Imports System.Web
Imports System.Drawing
' ExStart:MailMergingNamespace
Imports Aspose.Words.MailMerging
' ExEnd:MailMergingNamespace
Public Class MailMergeAlternatingRows
    Public Shared Sub Run()
        ' ExStart:MailMergeAlternatingRows           
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_MailMergeAndReporting()


        Dim doc As New Document(dataDir & Convert.ToString("MailMerge.AlternatingRows.doc"))

        ' Add a handler for the MergeField event.
        doc.MailMerge.FieldMergingCallback = New HandleMergeFieldAlternatingRows()

        ' Execute mail merge with regions.
        Dim dataTable As DataTable = GetSuppliersDataTable()
        doc.MailMerge.ExecuteWithRegions(dataTable)
        dataDir = dataDir & Convert.ToString("MailMerge.AlternatingRows_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:MailMergeAlternatingRows
        Console.WriteLine(Convert.ToString(vbLf & "Mail merge alternative rows performed successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
    ' ExStart:HandleMergeFieldAlternatingRows
    Private Class HandleMergeFieldAlternatingRows
        Implements IFieldMergingCallback
        ''' <summary>
        ''' Called for every merge field encountered in the document.
        ''' We can either return some data to the mail merge engine or do something
        ''' else with the document. In this case we modify cell formatting.
        ''' </summary>
        Private Sub IFieldMergingCallback_FieldMerging(ByVal e As FieldMergingArgs) Implements IFieldMergingCallback.FieldMerging
            If mBuilder Is Nothing Then
                mBuilder = New DocumentBuilder(e.Document)
            End If

            ' This way we catch the beginning of a new row.
            If e.FieldName.Equals("CompanyName") Then
                ' Select the color depending on whether the row number is even or odd.
                Dim rowColor As Color
                If IsOdd(mRowIdx) Then
                    rowColor = Color.FromArgb(213, 227, 235)
                Else
                    rowColor = Color.FromArgb(242, 242, 242)
                End If

                ' There is no way to set cell properties for the whole row at the moment,
                ' so we have to iterate over all cells in the row.
                For colIdx As Integer = 0 To 3
                    mBuilder.MoveToCell(0, mRowIdx, colIdx, 0)
                    mBuilder.CellFormat.Shading.BackgroundPatternColor = rowColor
                Next colIdx

                mRowIdx += 1
            End If
        End Sub

        Private Sub ImageFieldMerging(ByVal args As ImageFieldMergingArgs) Implements IFieldMergingCallback.ImageFieldMerging
            ' Do nothing.
        End Sub

        Private mBuilder As DocumentBuilder
        Private mRowIdx As Integer
    End Class
    ''' <summary>
    ''' Returns true if the value is odd; false if the value is even.
    ''' </summary>
    Private Shared Function IsOdd(value As Integer) As Boolean
        ' The code is a bit complex, but otherwise automatic conversion to VB does not work.
        Return ((value / 2) * 2).Equals(value)
    End Function
    ''' <summary>
    ''' Create DataTable and fill it with data.
    ''' In real life this DataTable should be filled from a database.
    ''' </summary>
    Private Shared Function GetSuppliersDataTable() As DataTable
        Dim dataTable As New DataTable("Suppliers")
        dataTable.Columns.Add("CompanyName")
        dataTable.Columns.Add("ContactName")
        For i As Integer = 0 To 9
            Dim datarow As DataRow = dataTable.NewRow()
            dataTable.Rows.Add(datarow)
            datarow(0) = "Company " + i.ToString()
            datarow(1) = "Contact " + i.ToString()
        Next
        Return dataTable
    End Function
    ' ExEnd:HandleMergeFieldAlternatingRows

End Class
