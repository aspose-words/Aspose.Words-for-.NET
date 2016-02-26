' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Public Class FieldsHelper
    Inherits DocumentVisitor
    ''' <summary>
    ''' Converts any fields of the specified type found in the descendants of the node into static text.
    ''' </summary>
    ''' <param name="compositeNode">The node in which all descendants of the specified FieldType will be converted to static text.</param>
    ''' <param name="targetFieldType">The FieldType of the field to convert to static text.</param>
    Public Shared Sub ConvertFieldsToStaticText(ByVal compositeNode As CompositeNode, ByVal targetFieldType As FieldType)
        Dim originalNodeText As String = compositeNode.ToString(SaveFormat.Text) 'ExSkip
        Dim helper As New FieldsHelper(targetFieldType)
        compositeNode.Accept(helper)

        Debug.Assert(originalNodeText.Equals(compositeNode.ToString(SaveFormat.Text)), "Error: Text of the node converted differs from the original") 'ExSkip
        For Each node As Node In compositeNode.GetChildNodes(NodeType.Any, True) 'ExSkip
            Debug.Assert(Not (TypeOf node Is FieldChar AndAlso (CType(node, FieldChar)).FieldType.Equals(targetFieldType)), "Error: A field node that should be removed still remains.") 'ExSkip
        Next node
    End Sub

    Private Sub New(ByVal targetFieldType As FieldType)
        mTargetFieldType = targetFieldType
    End Sub

    Public Overrides Function VisitFieldStart(ByVal fieldStart As FieldStart) As VisitorAction
        ' We must keep track of the starts and ends of fields incase of any nested fields.
        If fieldStart.FieldType.Equals(mTargetFieldType) Then
            mFieldDepth += 1
            fieldStart.Remove()
        Else
            ' This removes the field start if it's inside a field that is being converted.
            CheckDepthAndRemoveNode(fieldStart)
        End If

        Return VisitorAction.Continue
    End Function

    Public Overrides Function VisitFieldSeparator(ByVal fieldSeparator As FieldSeparator) As VisitorAction
        ' When visiting a field separator we should decrease the depth level.
        If fieldSeparator.FieldType.Equals(mTargetFieldType) Then
            mFieldDepth -= 1
            fieldSeparator.Remove()
        Else
            ' This removes the field separator if it's inside a field that is being converted.
            CheckDepthAndRemoveNode(fieldSeparator)
        End If

        Return VisitorAction.Continue
    End Function

    Public Overrides Function VisitFieldEnd(ByVal fieldEnd As FieldEnd) As VisitorAction
        If fieldEnd.FieldType.Equals(mTargetFieldType) Then
            fieldEnd.Remove()
        Else
            CheckDepthAndRemoveNode(fieldEnd) ' This removes the field end if it's inside a field that is being converted.
        End If

        Return VisitorAction.Continue
    End Function

    Public Overrides Function VisitRun(ByVal run As Run) As VisitorAction
        ' Remove the run if it is between the FieldStart and FieldSeparator of the field being converted.
        CheckDepthAndRemoveNode(run)

        Return VisitorAction.Continue
    End Function

    Public Overrides Function VisitParagraphEnd(ByVal paragraph As Paragraph) As VisitorAction
        If mFieldDepth > 0 Then
            ' The field code that is being converted continues onto another paragraph. We 
            ' need to copy the remaining content from this paragraph onto the next paragraph.
            Dim nextParagraph As Node = paragraph.NextSibling

            ' Skip ahead to the next available paragraph.
            Do While nextParagraph IsNot Nothing AndAlso nextParagraph.NodeType <> NodeType.Paragraph
                nextParagraph = nextParagraph.NextSibling
            Loop

            ' Copy all of the nodes over. Keep a list of these nodes so we know not to remove them.
            Do While paragraph.HasChildNodes
                mNodesToSkip.Add(paragraph.LastChild)
                CType(nextParagraph, Paragraph).PrependChild(paragraph.LastChild)
            Loop

            paragraph.Remove()
        End If

        Return VisitorAction.Continue
    End Function

    Public Overrides Function VisitTableStart(ByVal table As Table) As VisitorAction
        CheckDepthAndRemoveNode(table)

        Return VisitorAction.Continue
    End Function

    ''' <summary>
    ''' Checks whether the node is inside a field or should be skipped and then removes it if necessary.
    ''' </summary>
    Private Sub CheckDepthAndRemoveNode(ByVal node As Node)
        If mFieldDepth > 0 AndAlso (Not mNodesToSkip.Contains(node)) Then
            node.Remove()
        End If
    End Sub

    Private mFieldDepth As Integer = 0
    Private mNodesToSkip As New ArrayList()
    Private mTargetFieldType As FieldType
End Class
