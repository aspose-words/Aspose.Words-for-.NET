Imports Microsoft.VisualBasic
Imports System
Imports System.Reflection
Imports System.Collections
Imports System.IO
Imports System.Text

Imports Aspose.Words.Lists
Imports Aspose.Words.Fields
Imports Aspose.Words

Public Class ListUseDestinationStyles
    Public Shared Sub Run()
        ' ExStart:ListUseDestinationStyles
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_JoiningAndAppending()
        Dim fileName As String = "TestFile.Destination.doc"

        Dim dstDoc As New Document(dataDir & fileName)
        Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

        ' Set the source document to continue straight after the end of the destination document.
        srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous

        ' Keep track of the lists that are created.
        Dim newLists As New Hashtable()

        ' Iterate through all paragraphs in the document.
        For Each para As Paragraph In srcDoc.GetChildNodes(NodeType.Paragraph, True)
            If para.IsListItem Then
                Dim listId As Integer = para.ListFormat.List.ListId

                ' Check if the destination document contains a list with this ID already. If it does then this may
                ' cause the two lists to run together. Create a copy of the list in the source document instead.
                If dstDoc.Lists.GetListByListId(listId) IsNot Nothing Then
                    Dim currentList As List
                    ' A newly copied list already exists for this ID, retrieve the stored list and use it on 
                    ' the current paragraph.
                    If newLists.Contains(listId) Then
                        currentList = CType(newLists(listId), List)
                    Else
                        ' Add a copy of this list to the document and store it for later reference.
                        currentList = srcDoc.Lists.AddCopy(para.ListFormat.List)
                        newLists.Add(listId, currentList)
                    End If

                    ' Set the list of this paragraph  to the copied list.
                    para.ListFormat.List = currentList
                End If
            End If
        Next para

        ' Append the source document to end of the destination document.
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles)

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        ' Save the combined document to disk.
        dstDoc.Save(dataDir)
        ' ExEnd:ListUseDestinationStyles
        Console.WriteLine(vbNewLine & "Document appended successfully with list using destination styles." & vbNewLine & "File saved at " + dataDir)
    End Sub
End Class
