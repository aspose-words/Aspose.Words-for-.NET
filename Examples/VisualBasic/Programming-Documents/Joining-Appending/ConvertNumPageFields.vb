Imports Microsoft.VisualBasic
Imports System
Imports System.Reflection
Imports System.Collections
Imports System.IO
Imports System.Text

Imports Aspose.Words.Lists
Imports Aspose.Words.Fields
Imports Aspose.Words

Public Class ConvertNumPageFields
    Public Shared Sub Run()
        ' ExStart:ConvertNumPageFields
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_JoiningAndAppending()
        Dim fileName As String = "TestFile.Destination.doc"

        Dim dstDoc As New Document(dataDir & fileName)
        Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

        ' Restart the page numbering on the start of the source document.
        srcDoc.FirstSection.PageSetup.RestartPageNumbering = True
        srcDoc.FirstSection.PageSetup.PageStartingNumber = 1

        ' Append the source document to the end of the destination document.
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting)

        ' After joining the documents the NUMPAGE fields will now display the total number of pages which 
        ' Is undesired behavior. Call this method to fix them by replacing them with PAGEREF fields.
        ConvertNumPageFieldsToPageRef(dstDoc)

        ' This needs to be called in order to update the new fields with page numbers.
        dstDoc.UpdatePageLayout()

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        dstDoc.Save(dataDir)
        ' ExEnd:ConvertNumPageFields
        Console.WriteLine(vbNewLine & "Document appended successfully with converted NUMPAGE fields." & vbNewLine & "File saved at " + dataDir)
    End Sub
    ' ExStart:ConvertNumPageFieldsToPageRef
    ''' <summary>
    ''' Replaces all NUMPAGES fields in the document with PAGEREF fields. The replacement field displays the total number
    ''' Of pages in the sub document instead of the total pages in the document.
    ''' </summary>
    ''' <param name="doc">The combined document to process</param>
    Public Shared Sub ConvertNumPageFieldsToPageRef(ByVal doc As Document)
        ' This is the prefix for each bookmark which signals where page numbering restarts.
        ' The underscore "_" at the start inserts this bookmark as hidden in MS Word.
        Const bookmarkPrefix As String = "_SubDocumentEnd"
        ' Field name of the NUMPAGES field.
        Const numPagesFieldName As String = "NUMPAGES"
        ' Field name of the PAGEREF field.
        Const pageRefFieldName As String = "PAGEREF"

        ' Create a new DocumentBuilder which is used to insert the bookmarks and replacement fields.
        Dim builder As New DocumentBuilder(doc)
        ' Defines the number of page restarts that have been encountered and therefore the number of "sub" documents
        ' Found within this document.
        Dim subDocumentCount As Integer = 0

        ' Iterate through all sections in the document.
        For Each section As Section In doc.Sections
            ' This section has it' S page numbering restarted so we will treat this as the start of a sub document.
            ' Any PAGENUM fields in this inner document must be converted to special PAGEREF fields to correct numbering.
            If section.PageSetup.RestartPageNumbering Then
                ' Don' T do anything if this is the first section in the document. This part of the code will insert the bookmark marking
                ' The end of the previous sub document so therefore it is not applicable for first section in the document.
                If (Not section.Equals(doc.FirstSection)) Then
                    ' Get the previous section and the last node within the body of that section.
                    Dim prevSection As Section = CType(section.PreviousSibling, Section)
                    Dim lastNode As Node = prevSection.Body.LastChild

                    ' Use the DocumentBuilder to move to this node and insert the bookmark there.
                    ' This bookmark represents the end of the sub document.
                    builder.MoveTo(lastNode)
                    builder.StartBookmark(bookmarkPrefix & subDocumentCount)
                    builder.EndBookmark(bookmarkPrefix & subDocumentCount)

                    ' Increase the subdocument count to insert the correct bookmarks.
                    subDocumentCount += 1
                End If
            End If

            ' The last section simply needs the ending bookmark to signal that it is the end of the current sub document.
            If section.Equals(doc.LastSection) Then
                ' Insert the bookmark at the end of the body of the last section.
                ' Don' T increase the count this time as we are just marking the end of the document.
                Dim lastNode As Node = doc.LastSection.Body.LastChild
                builder.MoveTo(lastNode)
                builder.StartBookmark(bookmarkPrefix & subDocumentCount)
                builder.EndBookmark(bookmarkPrefix & subDocumentCount)
            End If

            ' Iterate through each NUMPAGES field in the section and replace the field with a PAGEREF field referring to the bookmark of the current subdocument
            ' This bookmark is positioned at the end of the sub document but does not exist yet. It is inserted when a section with restart page numbering or the last 
            ' Section is encountered.
            Dim nodes() As Node = section.GetChildNodes(NodeType.FieldStart, True).ToArray()
            For Each fieldStart As FieldStart In nodes
                If fieldStart.FieldType = FieldType.FieldNumPages Then
                    ' Get the field code.
                    Dim fieldCode As String = GetFieldCode(fieldStart)
                    ' Since the NUMPAGES field does not take any additional parameters we can assume the remaining part of the field
                    ' Code after the fieldname are the switches. We will use these to help recreate the NUMPAGES field as a PAGEREF field.
                    Dim fieldSwitches As String = fieldCode.Replace(numPagesFieldName, "").Trim()

                    ' Inserting the new field directly at the FieldStart node of the original field will cause the new field to
                    ' Not pick up the formatting of the original field. To counter this insert the field just before the original field
                    Dim previousNode As Node = fieldStart.PreviousSibling

                    ' If a previous run cannot be found then we are forced to use the FieldStart node.
                    If previousNode Is Nothing Then
                        previousNode = fieldStart
                    End If

                    ' Insert a PAGEREF field at the same position as the field.
                    builder.MoveTo(previousNode)
                    ' This will insert a new field with a code like " PAGEREF _SubDocumentEnd0 *\MERGEFORMAT ".
                    Dim newField As Field = builder.InsertField(String.Format(" {0} {1}{2} {3} ", pageRefFieldName, bookmarkPrefix, subDocumentCount, fieldSwitches))

                    ' The field will be inserted before the referenced node. Move the node before the field instead.
                    previousNode.ParentNode.InsertBefore(previousNode, newField.Start)

                    ' Remove the original NUMPAGES field from the document.
                    RemoveField(fieldStart)
                End If
            Next fieldStart
        Next section
    End Sub
    ' ExEnd:ConvertNumPageFieldsToPageRef
    ' ExStart:GetRemoveField
    ''' <summary>
    ''' Removes the Field from the document
    ''' </summary>
    ''' <param name="fieldStart">The field start node of the field to remove.</param>
    Private Shared Sub RemoveField(ByVal fieldStart As FieldStart)
        Dim currentNode As Node = fieldStart
        Dim isRemoving As Boolean = True
        Do While currentNode IsNot Nothing AndAlso isRemoving
            If currentNode.NodeType = NodeType.FieldEnd Then
                isRemoving = False
            End If

            Dim nextNode As Node = currentNode.NextPreOrder(currentNode.Document)
            currentNode.Remove()
            currentNode = nextNode
        Loop
    End Sub

    ''' <summary>
    ''' Retrieves the field code from a field.
    ''' </summary>
    ''' <param name="fieldStart">The field start of the field which to gather the field code from</param>
    ''' <returns></returns>
    Private Shared Function GetFieldCode(ByVal fieldStart As FieldStart) As String
        Dim builder As New StringBuilder()

        Dim node As Node = fieldStart
        Do While node IsNot Nothing AndAlso node.NodeType <> NodeType.FieldSeparator AndAlso node.NodeType <> NodeType.FieldEnd
            ' Use text only of Run nodes to avoid duplication.
            If node.NodeType = NodeType.Run Then
                builder.Append(node.GetText())
            End If
            node = node.NextPreOrder(node.Document)
        Loop
        Return builder.ToString()
    End Function
    ' ExEnd:GetRemoveField
End Class
