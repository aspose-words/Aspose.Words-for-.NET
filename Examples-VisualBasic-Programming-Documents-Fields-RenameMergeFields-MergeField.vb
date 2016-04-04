' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
''' <summary>
''' Represents a facade object for a merge field in a Microsoft Word document.
''' </summary>
Friend Class MergeField
    Friend Sub New(fieldStart As FieldStart)
        If fieldStart.Equals(Nothing) Then
            Throw New ArgumentNullException("fieldStart")
        End If
        If Not fieldStart.FieldType.Equals(FieldType.FieldMergeField) Then
            Throw New ArgumentException("Field start type must be FieldMergeField.")
        End If

        mFieldStart = fieldStart

        ' Find the field separator node.
        mFieldSeparator = fieldStart.GetField().Separator
        If mFieldSeparator Is Nothing Then
            Throw New InvalidOperationException("Cannot find field separator.")
        End If

        mFieldEnd = fieldStart.GetField().[End]
    End Sub

    ''' <summary>
    ''' Gets or sets the name of the merge field.
    ''' </summary>
    Friend Property Name() As String
        Get
            Return DirectCast(mFieldStart, FieldStart).GetField().Result.Replace("«", "").Replace("»", "")
        End Get
        Set(value As String)
            ' Merge field name is stored in the field result which is a Run
            ' node between field separator and field end.
            Dim fieldResult As Run = DirectCast(mFieldSeparator.NextSibling, Run)
            fieldResult.Text = String.Format("«{0}»", value)

            ' But sometimes the field result can consist of more than one run, delete these runs.
            RemoveSameParent(fieldResult.NextSibling, mFieldEnd)

            UpdateFieldCode(value)
        End Set
    End Property

    Private Sub UpdateFieldCode(fieldName As String)
        ' Field code is stored in a Run node between field start and field separator.
        Dim fieldCode As Run = DirectCast(mFieldStart.NextSibling, Run)

        Dim match As Match = gRegex.Match(DirectCast(mFieldStart, FieldStart).GetField().GetFieldCode())

        Dim newFieldCode As String = String.Format(" {0}{1} ", match.Groups("start").Value, fieldName)
        fieldCode.Text = newFieldCode

        ' But sometimes the field code can consist of more than one run, delete these runs.
        RemoveSameParent(fieldCode.NextSibling, mFieldSeparator)
    End Sub

    ''' <summary>
    ''' Removes nodes from start up to but not including the end node.
    ''' Start and end are assumed to have the same parent.
    ''' </summary>
    Private Shared Sub RemoveSameParent(ByVal startNode As Aspose.Words.Node, ByVal endNode As Aspose.Words.Node)
        If (endNode IsNot Nothing) AndAlso (startNode.ParentNode IsNot endNode.ParentNode) Then
            Throw New ArgumentException("Start and end nodes are expected to have the same parent.")
        End If

        Dim curChild As Aspose.Words.Node = startNode
        Do While (curChild IsNot Nothing) AndAlso (curChild IsNot endNode)
            Dim nextChild As Aspose.Words.Node = curChild.NextSibling
            curChild.Remove()
            curChild = nextChild
        Loop
    End Sub

    Private ReadOnly mFieldStart As Node
    Private ReadOnly mFieldSeparator As Node
    Private ReadOnly mFieldEnd As Node

    Private Shared ReadOnly gRegex As New Regex("\s*(?<start>MERGEFIELD\s|)(\s|)(?<name>\S+)\s+")
End Class
