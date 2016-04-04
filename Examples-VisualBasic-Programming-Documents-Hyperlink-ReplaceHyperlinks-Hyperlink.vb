' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
''' <summary>
''' This "facade" class makes it easier to work with a hyperlink field in a Word document.
'''
''' A hyperlink is represented by a HYPERLINK field in a Word document. A field in Aspose.Words
''' consists of several nodes and it might be difficult to work with all those nodes directly.
''' Note this is a simple implementation and will work only if the hyperlink code and name
''' each consist of one Run only.
'''
''' [FieldStart][Run - field code][FieldSeparator][Run - field result][FieldEnd]
'''
''' The field code contains a string in one of these formats:
''' HYPERLINK "url"
''' HYPERLINK \l "bookmark name"
'''
''' The field result contains text that is displayed to the user.
''' </summary>
Friend Class Hyperlink
    Friend Sub New(fieldStart As FieldStart)
        If fieldStart Is Nothing Then
            Throw New ArgumentNullException("fieldStart")
        End If
        If Not fieldStart.FieldType.Equals(FieldType.FieldHyperlink) Then
            Throw New ArgumentException("Field start type must be FieldHyperlink.")
        End If

        mFieldStart = fieldStart

        ' Find the field separator node.
        mFieldSeparator = fieldStart.GetField().Separator
        If mFieldSeparator Is Nothing Then
            Throw New InvalidOperationException("Cannot find field separator.")
        End If

        mFieldEnd = fieldStart.GetField().[End]

        ' Field code looks something like [ HYPERLINK "http:\\www.myurl.com" ], but it can consist of several runs.
        Dim fieldCode As String = fieldStart.GetField().GetFieldCode()
        Dim match As Match = gRegex.Match(fieldCode.Trim())
        mIsLocal = (match.Groups(1).Length > 0)
        'The link is local if \l is present in the field code.
        mTarget = match.Groups(2).Value
    End Sub

    ''' <summary>
    ''' Gets or sets the display name of the hyperlink.
    ''' </summary>
    Friend Property Name() As String
        Get
            Return GetTextSameParent(mFieldSeparator, mFieldEnd)
        End Get
        Set(value As String)
            ' Hyperlink display name is stored in the field result which is a Run
            ' node between field separator and field end.
            Dim fieldResult As Run = DirectCast(mFieldSeparator.NextSibling, Run)
            fieldResult.Text = value

            ' But sometimes the field result can consist of more than one run, delete these runs.
            RemoveSameParent(fieldResult.NextSibling, mFieldEnd)
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the target url or bookmark name of the hyperlink.
    ''' </summary>
    Friend Property Target() As String
        Get
            Dim dummy As String = Nothing
            ' This is needed to fool the C# to VB.NET converter.
            Return mTarget
        End Get
        Set(value As String)
            mTarget = value
            UpdateFieldCode()
        End Set
    End Property

    ''' <summary>
    ''' True if the hyperlink's target is a bookmark inside the document. False if the hyperlink is a url.
    ''' </summary>
    Friend Property IsLocal() As Boolean
        Get
            Return mIsLocal
        End Get
        Set(value As Boolean)
            mIsLocal = value
            UpdateFieldCode()
        End Set
    End Property

    Private Sub UpdateFieldCode()
        ' Field code is stored in a Run node between field start and field separator.
        Dim fieldCode As Run = DirectCast(mFieldStart.NextSibling, Run)
        fieldCode.Text = String.Format("HYPERLINK {0}""{1}""", (If((mIsLocal), "\l ", "")), mTarget)

        ' But sometimes the field code can consist of more than one run, delete these runs.
        RemoveSameParent(fieldCode.NextSibling, mFieldSeparator)
    End Sub

    ''' <summary>
    ''' Retrieves text from start up to but not including the end node.
    ''' </summary>
    Private Shared Function GetTextSameParent(startNode As Node, endNode As Node) As String
        If (endNode IsNot Nothing) AndAlso (startNode.ParentNode IsNot endNode.ParentNode) Then
            Throw New ArgumentException("Start and end nodes are expected to have the same parent.")
        End If

        Dim builder As New StringBuilder()
        Dim child As Node = startNode
        While Not child.Equals(endNode)
            builder.Append(child.GetText())
            child = child.NextSibling
        End While

        Return builder.ToString()
    End Function

    ''' <summary>
    ''' Removes nodes from start up to but not including the end node.
    ''' Start and end are assumed to have the same parent.
    ''' </summary>
    Private Shared Sub RemoveSameParent(ByVal startNode As Aspose.Words.Node, ByVal endNode As Aspose.Words.Node)
        If ((endNode IsNot Nothing) AndAlso (startNode IsNot Nothing)) AndAlso (startNode.ParentNode IsNot endNode.ParentNode) Then
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
    Private mIsLocal As Boolean
    Private mTarget As String

    
    Private Shared ReadOnly gRegex As New Regex("\S+" + "\s+" + "(?:""""\s+)?" + "(\\l\s+)?" + """" + "([^""]+)" + """")
End Class
