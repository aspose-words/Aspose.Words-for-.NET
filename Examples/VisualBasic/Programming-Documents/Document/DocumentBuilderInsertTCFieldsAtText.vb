Imports System.IO
Imports Aspose.Words
Imports System.Drawing
Imports Aspose.Words.Tables
Imports Aspose.Words.Replacing
Imports System.Text.RegularExpressions

Class DocumentBuilderInsertTCFieldsAtText
    Public Shared Sub Run()
        ' ExStart:DocumentBuilderInsertTCFieldsAtText
        Dim doc As New Document()

        Dim options As New FindReplaceOptions()
        options.ReplacingCallback = New InsertTCFieldHandler("Chapter 1", "\l 1")

        ' Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document.
        doc.Range.Replace(New Regex("The Beginning"), "", options)
        ' ExEnd:DocumentBuilderInsertTCFieldsAtText

    End Sub
End Class
' ExStart:InsertTCFieldHandler
Public Class InsertTCFieldHandler
    Implements IReplacingCallback
    ' Store the text and switches to be used for the TC fields.
    Private mFieldText As String
    Private mFieldSwitches As String

    ''' <summary>
    ''' The switches to use for each TC field. Can be an empty string or null.
    ''' </summary>
    Public Sub New(switches As String)
        Me.New(String.Empty, switches)
        mFieldSwitches = switches
    End Sub

    ''' <summary>
    ''' The display text and switches to use for each TC field. Display name can be an empty string or null.
    ''' </summary>
    Public Sub New(text As String, switches As String)
        mFieldText = text
        mFieldSwitches = switches
    End Sub

    Private Function IReplacingCallback_Replacing(args As ReplacingArgs) As ReplaceAction Implements IReplacingCallback.Replacing
        ' Create a builder to insert the field.
        Dim builder As New DocumentBuilder(DirectCast(args.MatchNode.Document, Aspose.Words.Document))
        ' Move to the first node of the match.
        builder.MoveTo(args.MatchNode)

        ' If the user specified text to be used in the field as display text then use that, otherwise use the 
        ' match string as the display text.
        Dim insertText As String

        If Not String.IsNullOrEmpty(mFieldText) Then
            insertText = mFieldText
        Else
            insertText = args.Match.Value
        End If

        ' Insert the TC field before this node using the specified string as the display text and user defined switches.
        builder.InsertField(String.Format("TC ""{0}"" {1}", insertText, mFieldSwitches))

        ' We have done what we want so skip replacement.
        Return ReplaceAction.Skip
    End Function
End Class
' ExEnd:InsertTCFieldHandler

