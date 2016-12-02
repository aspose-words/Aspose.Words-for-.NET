Imports System
Imports System.Text
Imports System.Text.RegularExpressions
Imports Aspose.Words
Imports Aspose.Words.Fields
Public Class RenameMergeFields
    Public Shared Sub Run()
        ' ExStart:RenameMergeFields
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()

        ' Specify your document name here.
        Dim doc As New Document(dataDir & Convert.ToString("RenameMergeFields.doc"))

        ' Select all field start nodes so we can find the merge fields.
        Dim fieldStarts As NodeCollection = doc.GetChildNodes(NodeType.FieldStart, True)
        For Each fieldStart As FieldStart In fieldStarts
            If fieldStart.FieldType.Equals(FieldType.FieldMergeField) Then
                Dim mergeField As New MergeField(fieldStart)
                mergeField.Name = mergeField.Name & Convert.ToString("_Renamed")
            End If
        Next

        dataDir = dataDir & Convert.ToString("RenameMergeFields_out.doc")
        doc.Save(dataDir)
        ' ExEnd:RenameMergeFields
        Console.WriteLine(Convert.ToString(vbLf & "Merge fields rename successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
    ' ExStart:MergeField
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
                ' Node between field separator and field end.
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
    ' ExEnd:MergeField

End Class
