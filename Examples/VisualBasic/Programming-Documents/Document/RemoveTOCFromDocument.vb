Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Fields
Imports System.Collections
Class RemoveTOCFromDocument
    Public Shared Sub Run()
        ' ExStart:RemoveTOCFromDocument
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithStyles()

        ' Open a document which contains a TOC.
        Dim doc As New Document(dataDir & Convert.ToString("Document.TableOfContents.doc"))

        ' Remove the first table of contents from the document.
        RemoveTableOfContents(doc, 0)

        dataDir = dataDir & Convert.ToString("Document.TableOfContentsRemoveToc_out_.doc")
        ' Save the output.
        doc.Save(dataDir)

        Console.WriteLine(Convert.ToString(vbLf & "Specified TOC from a document removed successfully." & vbLf & "File saved at ") & dataDir)
    End Sub

    ''' <summary>
    ''' Removes the specified table of contents field from the document.
    ''' </summary>
    ''' <param name="doc">The document to remove the field from.</param>
    ''' <param name="index">The zero-based index of the TOC to remove.</param>
    Public Shared Sub RemoveTableOfContents(doc As Document, index As Integer)
        ' Store the FieldStart nodes of TOC fields in the document for quick access.
        Dim fieldStarts As New ArrayList()
        ' This is a list to store the nodes found inside the specified TOC. They will be removed
        ' at the end of this method.
        Dim nodeList As New ArrayList()

        For Each start As FieldStart In doc.GetChildNodes(NodeType.FieldStart, True)
            If start.FieldType = FieldType.FieldTOC Then
                ' Add all FieldStarts which are of type FieldTOC.
                fieldStarts.Add(start)
            End If
        Next

        ' Ensure the TOC specified by the passed index exists.
        If index > fieldStarts.Count - 1 Then
            Throw New ArgumentOutOfRangeException("TOC index is out of range")
        End If

        Dim isRemoving As Boolean = True
        ' Get the FieldStart of the specified TOC.
        Dim currentNode As Node = DirectCast(fieldStarts(index), Node)

        While isRemoving
            ' It is safer to store these nodes and delete them all at once later.
            nodeList.Add(currentNode)
            currentNode = currentNode.NextPreOrder(doc)

            ' Once we encounter a FieldEnd node of type FieldTOC then we know we are at the end
            ' of the current TOC and we can stop here.
            If currentNode.NodeType = NodeType.FieldEnd Then
                Dim fieldEnd As FieldEnd = DirectCast(currentNode, FieldEnd)
                If fieldEnd.FieldType = FieldType.FieldTOC Then
                    isRemoving = False
                End If
            End If
        End While

        ' Remove all nodes found in the specified TOC.
        For Each node As Node In nodeList
            node.Remove()
        Next
    End Sub
    ' ExEnd:RemoveTOCFromDocument

End Class

