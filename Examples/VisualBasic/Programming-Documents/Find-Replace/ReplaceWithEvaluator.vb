Imports System.Text.RegularExpressions
Imports System.Collections
Imports System.Drawing
Imports System.IO
Imports System.Reflection
Imports Aspose.Words
Imports Aspose.Words.Replacing

Class ReplaceWithEvaluator
    Public Shared Sub Run()
        ' ExStart:ReplaceWithEvaluator
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_FindAndReplace()
        Dim doc As New Document(dataDir & Convert.ToString("Range.ReplaceWithEvaluator.doc"))
        Dim options As New FindReplaceOptions()
        options.ReplacingCallback = New MyReplaceEvaluator()

        doc.Range.Replace(New Regex("[s|m]ad"), "", options)

        dataDir = dataDir & Convert.ToString("Range.ReplaceWithEvaluator_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:ReplaceWithEvaluator
        Console.WriteLine(Convert.ToString(vbLf & "Text replaced successfully with evaluator." & vbLf & "File saved at ") & dataDir)
    End Sub
    ' ExStart:MyReplaceEvaluator
    Private Class MyReplaceEvaluator
        Implements IReplacingCallback
        ''' <summary>
        ''' This is called during a replace operation each time a match is found.
        ''' This method appends a number to the match string and returns it as a replacement string.
        ''' </summary>
        Private Function IReplacingCallback_Replacing(e As ReplacingArgs) As ReplaceAction Implements IReplacingCallback.Replacing
            e.Replacement = e.Match.ToString() + mMatchNumber.ToString()
            mMatchNumber += 1
            Return ReplaceAction.Replace
        End Function

        Private mMatchNumber As Integer
    End Class
    ' ExEnd:MyReplaceEvaluator        
End Class

