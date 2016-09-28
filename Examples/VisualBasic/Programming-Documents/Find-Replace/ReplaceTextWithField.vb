Imports Microsoft.VisualBasic
Imports System.Collections
Imports System.IO
Imports System.Text.RegularExpressions
Imports Aspose.Words
Imports Aspose.Words.Fields
Imports Aspose.Words.Replacing

Public Class ReplaceTextWithField
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_FindAndReplace()
        Dim fileName As String = "Field.ReplaceTextWithFields.doc"

        Dim doc As New Document(dataDir & fileName)
        Dim options As New FindReplaceOptions()
        options.ReplacingCallback = New ReplaceTextWithFieldHandler(FieldType.FieldMergeField)

        ' Replace any "PlaceHolderX" instances in the document (where X is a number) with a merge field.
        doc.Range.Replace(New Regex("PlaceHolder(\d+)"), "", options)

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)

        doc.Save(dataDir)

        Console.WriteLine(vbNewLine & "Text replaced with field successfully." & vbNewLine & "File saved at " + dataDir)
    End Sub
End Class

Public Class ReplaceTextWithFieldHandler
    Implements IReplacingCallback
    Public Sub New(ByVal type As FieldType)
        mFieldType = type
    End Sub

    Public Function Replacing(ByVal args As ReplacingArgs) As ReplaceAction Implements IReplacingCallback.Replacing
        Dim runs As ArrayList = FindAndSplitMatchRuns(args)

        ' Create DocumentBuilder which is used to insert the field.
        Dim builder As New DocumentBuilder(CType(args.MatchNode.Document, Document))
        builder.MoveTo(CType(runs(runs.Count - 1), Run))

        ' Calculate the name of the field from the FieldType enumeration by removing the first instance of "Field" from the text. 
        ' This works for almost all of the field types.
        Dim fieldName As String = mFieldType.ToString().ToUpper().Substring(5)

        ' Insert the field into the document using the specified field type and the match text as the field name.
        ' If the fields you are inserting do not require this extra parameter then it can be removed from the string below.
        builder.InsertField(String.Format("{0} {1}", fieldName, args.Match.Groups(0)))

        ' Now remove all runs in the sequence.
        For Each run As Run In runs
            run.Remove()
        Next run

        ' Signal to the replace engine to do nothing because we have already done all what we wanted.
        Return ReplaceAction.Skip
    End Function

    ''' <summary>
    ''' Finds and splits the match runs and returns them in an ArrayList.
    ''' </summary>
    Public Function FindAndSplitMatchRuns(ByVal args As ReplacingArgs) As ArrayList
        ' This is a Run node that contains either the beginning or the complete match.
        Dim currentNode As Node = args.MatchNode

        ' The first (and may be the only) run can contain text before the match, 
        ' in this case it is necessary to split the run.
        If args.MatchOffset > 0 Then
            currentNode = SplitRun(CType(currentNode, Run), args.MatchOffset)
        End If

        ' This array is used to store all nodes of the match for further removing.
        Dim runs As New ArrayList()

        ' Find all runs that contain parts of the match string.
        Dim remainingLength As Integer = args.Match.Value.Length
        Do While (remainingLength > 0) AndAlso (currentNode IsNot Nothing) AndAlso (currentNode.GetText().Length <= remainingLength)
            runs.Add(currentNode)
            remainingLength = remainingLength - currentNode.GetText().Length

            ' Select the next Run node. 
            ' Have to loop because there could be other nodes such as BookmarkStart etc.
            Do
                currentNode = currentNode.NextSibling
            Loop While (currentNode IsNot Nothing) AndAlso (currentNode.NodeType <> NodeType.Run)
        Loop

        ' Split the last run that contains the match if there is any text left.
        If (currentNode IsNot Nothing) AndAlso (remainingLength > 0) Then
            SplitRun(CType(currentNode, Run), remainingLength)
            runs.Add(currentNode)
        End If

        Return runs
    End Function

    ''' <summary>
    ''' Splits text of the specified run into two runs.
    ''' Inserts the new run just after the specified run.
    ''' </summary>
    Private Function SplitRun(ByVal run As Run, ByVal position As Integer) As Run
        Dim afterRun As Run = CType(run.Clone(True), Run)
        afterRun.Text = run.Text.Substring(position)
        run.Text = run.Text.Substring(0, position)
        run.ParentNode.InsertAfter(afterRun, run)
        Return afterRun

    End Function

    Private mFieldType As FieldType
End Class