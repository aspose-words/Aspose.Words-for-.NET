Imports Microsoft.VisualBasic
Imports System
Imports System.IO

Imports Aspose.Words

Public Class WorkingWithNodes
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_QuickStart()

        ' Create a new document.
        Dim doc As New Document()

        ' Creates and adds a paragraph node to the document.
        Dim para As New Paragraph(doc)

        ' Typed access to the last section of the document.
        Dim section As Section = doc.LastSection
        section.Body.AppendChild(para)

        ' Next print the node type of one of the nodes in the document.
        Dim nodeType As NodeType = doc.FirstSection.Body.NodeType

        Console.WriteLine("NodeType: " & Node.NodeTypeToString(nodeType))


    End Sub
End Class
