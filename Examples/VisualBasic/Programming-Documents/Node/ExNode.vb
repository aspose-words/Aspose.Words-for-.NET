Imports Aspose.Words
Imports Aspose.Words.Saving
Imports Aspose.Words.Tables

Public Class ExNode
    Public Shared Sub Run()
        ' The following method shows how to use the NodeType enumeration.
        UseNodeType()
        ' The following method shows how to access the parent node.
        GetParentNode()
        ' The following method shows that when you create any node, it requires a document that will own the node.
        OwnerDocument()
        ' Shows how to extract a specific child node from a CompositeNode by using the GetChild method and passing the NodeType and index.
        EnumerateChildNodes()
        ' Shows how to enumerate immediate children of a CompositeNode using indexed access.
        IndexChildNodes()
        ' Shows how to efficiently visit all direct and indirect children of a composite node.
        RecurseAllNodes()
        ' Demonstrates how to use typed properties to access nodes of the document tree.
        TypedAccess()
        ' The following method shows how to creates and adds a paragraph node.
        CreateAndAddParagraphNode()
    End Sub
    Public Shared Sub UseNodeType()
        ' ExStart:UseNodeType            
        Dim doc As New Document()
        ' Returns NodeType.Document
        Dim type As NodeType = doc.NodeType
        ' ExEnd:UseNodeType
    End Sub
    Public Shared Sub GetParentNode()
        ' ExStart:GetParentNode           
        ' Create a new empty document. It has one section.
        Dim doc As New Document()
        ' The section is the first child node of the document.
        Dim section As Aspose.Words.Node = doc.FirstChild
        ' The section's parent node is the document.
        Console.WriteLine("Section parent is the document: " & (doc Is section.ParentNode))
        ' ExEnd:GetParentNode           
    End Sub
    Public Shared Sub OwnerDocument()
        ' ExStart:OwnerDocument            
        ' Open a file from disk.
        Dim doc As New Document()

        ' Creating a new node of any type requires a document passed into the constructor.
        Dim para As New Paragraph(doc)

        ' The new paragraph node does not yet have a parent.
        Console.WriteLine("Paragraph has no parent node: " & (para.ParentNode Is Nothing))

        ' But the paragraph node knows its document.
        Console.WriteLine("Both nodes' documents are the same: " & (para.Document Is doc))

        ' The fact that a node always belongs to a document allows us to access and modify 
        ' properties that reference the document-wide data such as styles or lists.
        para.ParagraphFormat.StyleName = "Heading 1"

        ' Now add the paragraph to the main text of the first section.
        doc.FirstSection.Body.AppendChild(para)

        ' The paragraph node is now a child of the Body node.
        Console.WriteLine("Paragraph has a parent node: " & (para.ParentNode IsNot Nothing))
        ' ExEnd:OwnerDocument
    End Sub
    Public Shared Sub EnumerateChildNodes()
        ' ExStart:EnumerateChildNodes 
        Dim doc As New Document()
        Dim paragraph As Paragraph = DirectCast(doc.GetChild(NodeType.Paragraph, 0, True), Paragraph)

        Dim children As NodeCollection = paragraph.ChildNodes
        For Each child As Aspose.Words.Node In children
            ' Paragraph may contain children of various types such as runs, shapes and so on.
            If child.NodeType.Equals(NodeType.Run) Then
                ' Say we found the node that we want, do something useful.
                Dim run As Run = DirectCast(child, Run)
                Console.WriteLine(run.Text)
            End If
        Next
        ' ExEnd:EnumerateChildNodes
    End Sub
    Public Shared Sub IndexChildNodes()
        ' ExStart:IndexChildNodes
        Dim doc As New Document()
        Dim paragraph As Paragraph = DirectCast(doc.GetChild(NodeType.Paragraph, 0, True), Paragraph)
        Dim children As NodeCollection = paragraph.ChildNodes
        For i As Integer = 0 To children.Count - 1
            Dim child As Aspose.Words.Node = children(i)

            ' Paragraph may contain children of various types such as runs, shapes and so on.
            If child.NodeType.Equals(NodeType.Run) Then
                ' Say we found the node that we want, do something useful.
                Dim run As Run = DirectCast(child, Run)
                Console.WriteLine(run.Text)
            End If
        Next
        ' ExEnd:IndexChildNodes
    End Sub
    ' ExStart:RecurseAllNodes
    Public Shared Sub RecurseAllNodes()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithNode()
        ' Open a document.
        Dim doc As New Document(dataDir & Convert.ToString("Node.RecurseAllNodes.doc"))

        ' Invoke the recursive function that will walk the tree.
        TraverseAllNodes(doc)
    End Sub

    ''' <summary>
    ''' A simple function that will walk through all children of a specified node recursively 
    ''' and print the type of each node to the screen.
    ''' </summary>
    Public Shared Sub TraverseAllNodes(parentNode As CompositeNode)
        ' This is the most efficient way to loop through immediate children of a node.
        Dim childNode As Aspose.Words.Node = parentNode.FirstChild
        While childNode IsNot Nothing
            ' Do some useful work.
            Console.WriteLine(Aspose.Words.Node.NodeTypeToString(childNode.NodeType))

            ' Recurse into the node if it is a composite node.
            If childNode.IsComposite Then
                TraverseAllNodes(DirectCast(childNode, CompositeNode))
            End If
            childNode = childNode.NextSibling
        End While
    End Sub
    ' ExEnd:RecurseAllNodes
    Public Shared Sub TypedAccess()
        ' ExStart:TypedAccess
        Dim doc As New Document()
        Dim section As Section = doc.FirstSection
        ' Quick typed access to the Body child node of the Section.
        Dim body As Body = section.Body
        ' Quick typed access to all Table child nodes contained in the Body.
        Dim tables As TableCollection = body.Tables

        For Each table As Table In tables
            ' Quick typed access to the first row of the table.
            If table.FirstRow IsNot Nothing Then
                table.FirstRow.Remove()
            End If

            ' Quick typed access to the last row of the table.
            If table.LastRow IsNot Nothing Then
                table.LastRow.Remove()
            End If
        Next
        ' ExEnd:TypedAccess
    End Sub

    Public Shared Sub CreateAndAddParagraphNode()
        ' ExStart:CreateAndAddParagraphNode
        Dim doc As New Document()
        Dim para As New Paragraph(doc)
        Dim section As Aspose.Words.Section = doc.LastSection
        section.Body.AppendChild(para)
        ' ExEnd:CreateAndAddParagraphNode
    End Sub
End Class
