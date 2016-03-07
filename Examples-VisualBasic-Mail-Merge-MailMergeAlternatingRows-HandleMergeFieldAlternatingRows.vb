' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
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
