' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System

Imports Aspose.Words
Imports Aspose.Words.Saving
Imports Aspose.Words.Tables

Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Public Class ExNode
		Inherits ApiExampleBase
		<Test> _
		Public Sub UseNodeType()
			'ExStart
			'ExFor:NodeType
			'ExId:UseNodeType
			'ExSummary:The following example shows how to use the NodeType enumeration.
			Dim doc As New Document()

			' Returns NodeType.Document
			Dim type As NodeType = doc.NodeType
			'ExEnd
		End Sub

		<Test> _
		Public Sub CloneCompositeNode()
			'ExStart
			'ExFor:Node
			'ExFor:Node.Clone
			'ExSummary:Shows how to clone composite nodes with and without their child nodes.
			' Create a new empty document.
			Dim doc As New Document()

			' Add some text to the first paragraph
			Dim para As Paragraph = doc.FirstSection.Body.FirstParagraph
			para.AppendChild(New Run(doc, "Some text"))

			' Clone the paragraph and the child nodes.
			Dim cloneWithChildren As Node = para.Clone(True)
			' Only clone the paragraph and no child nodes.
			Dim cloneWithoutChildren As Node = para.Clone(False)
			'ExEnd

			Assert.IsTrue((CType(cloneWithChildren, CompositeNode)).HasChildNodes)
			Assert.IsFalse((CType(cloneWithoutChildren, CompositeNode)).HasChildNodes)
		End Sub

		<Test> _
		Public Sub GetParentNode()
			'ExStart
			'ExFor:Node.ParentNode
			'ExId:AccessParentNode
			'ExSummary:Shows how to access the parent node.
			' Create a new empty document. It has one section.
			Dim doc As New Document()

			' The section is the first child node of the document.
			Dim section As Node = doc.FirstChild

			' The section's parent node is the document.
			Console.WriteLine("Section parent is the document: " & (doc Is section.ParentNode))
			'ExEnd

			Assert.AreEqual(doc, section.ParentNode)
		End Sub

		<Test> _
		Public Sub OwnerDocument()
			'ExStart
			'ExFor:Node.Document
			'ExFor:Node.ParentNode
			'ExId:CreatingNodeRequiresDocument
			'ExSummary:Shows that when you create any node, it requires a document that will own the node.
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
			'ExEnd

			Assert.AreEqual(doc, para.Document)
			Assert.IsNotNull(para.ParentNode)
		End Sub

		<Test> _
		Public Sub EnumerateChildNodes()
			Dim doc As New Document()
			'ExStart
			'ExFor:Node
			'ExFor:CompositeNode
			'ExFor:CompositeNode.GetChild
			'ExSummary:Shows how to extract a specific child node from a CompositeNode by using the GetChild method and passing the NodeType and index.
			Dim paragraph As Paragraph = CType(doc.GetChild(NodeType.Paragraph, 0, True), Paragraph)
			'ExEnd

			'ExStart
			'ExFor:CompositeNode.ChildNodes
			'ExFor:CompositeNode.GetEnumerator
			'ExId:ChildNodesForEach
			'ExSummary:Shows how to enumerate immediate children of a CompositeNode using the enumerator provided by the ChildNodes collection.
			Dim children As NodeCollection = paragraph.ChildNodes
			For Each child As Node In children
				' Paragraph may contain children of various types such as runs, shapes and so on.
				If child.NodeType.Equals(NodeType.Run) Then
					' Say we found the node that we want, do something useful.
					Dim run As Run = CType(child, Run)
					Console.WriteLine(run.Text)
				End If
			Next child
			'ExEnd
		End Sub

		<Test> _
		Public Sub IndexChildNodes()
			Dim doc As New Document()
			Dim paragraph As Paragraph = CType(doc.GetChild(NodeType.Paragraph, 0, True), Paragraph)

			'ExStart
			'ExFor:NodeCollection.Count
			'ExFor:NodeCollection.Item
			'ExId:ChildNodesIndexer
			'ExSummary:Shows how to enumerate immediate children of a CompositeNode using indexed access.
			Dim children As NodeCollection = paragraph.ChildNodes
			For i As Integer = 0 To children.Count - 1
				Dim child As Node = children(i)

				' Paragraph may contain children of various types such as runs, shapes and so on.
				If child.NodeType.Equals(NodeType.Run) Then
					' Say we found the node that we want, do something useful.
					Dim run As Run = CType(child, Run)
					Console.WriteLine(run.Text)
				End If
			Next i
			'ExEnd
		End Sub

		''' <summary>
		''' This calls the below method to resolve skipping of [Test] in VB.NET.
		''' </summary>
		<Test> _
		Public Sub RecurseAllNodesCaller()
			Me.RecurseAllNodes()
		End Sub

		'ExStart
		'ExFor:Node.NextSibling
		'ExFor:CompositeNode.FirstChild
		'ExFor:Node.IsComposite
		'ExFor:CompositeNode.IsComposite
		'ExFor:Node.NodeTypeToString
		'ExId:RecurseAllNodes            
		'ExSummary:Shows how to efficiently visit all direct and indirect children of a composite node.
		Public Sub RecurseAllNodes()
			' Open a document.
			Dim doc As New Document(MyDir & "Node.RecurseAllNodes.doc")

			' Invoke the recursive function that will walk the tree.
			Me.TraverseAllNodes(doc)
		End Sub

		''' <summary>
		''' A simple function that will walk through all children of a specified node recursively 
		''' and print the type of each node to the screen.
		''' </summary>
		Public Sub TraverseAllNodes(ByVal parentNode As CompositeNode)
			' This is the most efficient way to loop through immediate children of a node.
			Dim childNode As Node = parentNode.FirstChild
			Do While childNode IsNot Nothing
				' Do some useful work.
				Console.WriteLine(Node.NodeTypeToString(childNode.NodeType))

				' Recurse into the node if it is a composite node.
				If childNode.IsComposite Then
					Me.TraverseAllNodes(CType(childNode, CompositeNode))
				End If
				childNode = childNode.NextSibling
			Loop
		End Sub
		'ExEnd


		<Test> _
		Public Sub RemoveNodes()
			Dim doc As New Document()

			'ExStart
			'ExFor:Node
			'ExFor:Node.NodeType
			'ExFor:Node.Remove
			'ExSummary:Shows how to remove all nodes of a specific type from a composite node. In this example we remove tables from a section body.
			' Get the section that we want to work on.
			Dim section As Section = doc.Sections(0)
			Dim body As Body = section.Body

			' Select the first child node in the body.
			Dim curNode As Node = body.FirstChild

			Do While curNode IsNot Nothing
				' Save the pointer to the next sibling node because if the current 
				' node is removed from the parent in the next step, we will have 
				' no way of finding the next node to continue the loop.
				Dim nextNode As Node = curNode.NextSibling

				' A section body can contain Paragraph and Table nodes.
				' If the node is a Table, remove it from the parent.
				If curNode.NodeType.Equals(NodeType.Table) Then
					curNode.Remove()
				End If

				' Continue going through child nodes until null (no more siblings) is reached.
				curNode = nextNode
			Loop
			'ExEnd
		End Sub

		<Test> _
		Public Sub EnumNextSibling()
			Dim doc As New Document()

			'ExStart
			'ExFor:CompositeNode.FirstChild
			'ExFor:Node.NextSibling
			'ExFor:Node.NodeTypeToString
			'ExFor:Node.NodeType
			'ExSummary:Shows how to enumerate immediate child nodes of a composite node using NextSibling. In this example we enumerate all paragraphs of a section body.
			' Get the section that we want to work on.
			Dim section As Section = doc.Sections(0)
			Dim body As Body = section.Body

			' Loop starting from the first child until we reach null.
			Dim node As Node = body.FirstChild
			Do While node IsNot Nothing
				' Output the types of the nodes that we come across.
				Console.WriteLine(Node.NodeTypeToString(node.NodeType))
				node = node.NextSibling
			Loop
			'ExEnd
		End Sub

		<Test> _
		Public Sub TypedAccess()
			Dim doc As New Document()

			'ExStart
			'ExFor:Story.Tables
			'ExFor:Table.FirstRow
			'ExFor:Table.LastRow
			'ExFor:TableCollection
			'ExId:TypedPropertiesAccess
			'ExSummary:Demonstrates how to use typed properties to access nodes of the document tree.
			' Quick typed access to the first child Section node of the Document.
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
			Next table
			'ExEnd
		End Sub

		<Test> _
		Public Sub UpdateFieldsInRange()
			Dim doc As New Document()

			'ExStart
			'ExFor:Range.UpdateFields
			'ExSummary:Demonstrates how to update document fields in the body of the first section only.
			doc.FirstSection.Body.Range.UpdateFields()
			'ExEnd
		End Sub

		<Test> _
		Public Sub RemoveChild()
			Dim doc As New Document()

			'ExStart
			'ExFor:CompositeNode.LastChild
			'ExFor:Node.PreviousSibling
			'ExFor:CompositeNode.RemoveChild
			'ExSummary:Demonstrates use of methods of Node and CompositeNode to remove a section before the last section in the document.
			' Document is a CompositeNode and LastChild returns the last child node in the Document node.
			' Since the Document can contain only Section nodes, the last child is the last section.
			Dim lastSection As Node = doc.LastChild

			' Each node knows its next and previous sibling nodes.
			' Previous sibling of a section is a section before the specified section.
			' If the node is the first child, PreviousSibling will return null.
			Dim sectionBeforeLast As Node = lastSection.PreviousSibling

			If sectionBeforeLast IsNot Nothing Then
				doc.RemoveChild(sectionBeforeLast)
			End If
			'ExEnd
		End Sub

		<Test> _
		Public Sub CompositeNode_SelectNodes()
			'ExStart
			'ExFor:CompositeNode.SelectSingleNode
			'ExFor:CompositeNode.SelectNodes
			'ExSummary:Shows how to select certain nodes by using an XPath expression.
			Dim doc As New Document(MyDir & "Table.Document.doc")

			' This expression will extract all paragraph nodes which are descendants of any table node in the document.
			' This will return any paragraphs which are in a table.
			Dim nodeList As NodeList = doc.SelectNodes("//Table//Paragraph")

			' This expression will select any paragraphs that are direct children of any body node in the document.
			nodeList = doc.SelectNodes("//Body/Paragraph")

			' Use SelectSingleNode to select the first result of the same expression as above.
			Dim node As Node = doc.SelectSingleNode("//Body/Paragraph")
			'ExEnd
		End Sub

		<Test> _
		Public Sub TestNodeIsInsideField()
			'ExStart:
			'ExFor:CompositeNode.SelectNodes
			'ExFor:CompositeNode.GetChild
			'ExSummary:Shows how to test if a node is inside a field by using an XPath expression.
			' Let's pick a document we know has some fields in.
			Dim doc As New Document(MyDir & "MailMerge.MergeImage.doc")

			' Let's say we want to check if the Run below is inside a field.
			Dim run As Run = CType(doc.GetChild(NodeType.Run, 5, True), Run)

			' Evaluate the XPath expression. The resulting NodeList will contain all nodes found inside a field a field (between FieldStart 
			' and FieldEnd exclusive). There can however be FieldStart and FieldEnd nodes in the list if there are nested fields 
			' in the path. Currently does not find rare fields in which the FieldCode or FieldResult spans across multiple paragraphs.
			Dim resultList As NodeList = doc.SelectNodes("//FieldStart/following-sibling::node()[following-sibling::FieldEnd]")

			' Check if the specified run is one of the nodes that are inside the field.
			For Each node As Node In resultList
				If node Is run Then
					Console.WriteLine("The node is found inside a field")
					Exit For
				End If
			Next node
			'ExEnd
		End Sub

		<Test> _
		Public Sub CreateAndAddParagraphNode()
			'ExStart
			'ExId:CreateAndAddParagraphNode
			'ExSummary:Creates and adds a paragraph node.
			Dim doc As New Document()

			Dim para As New Paragraph(doc)

			Dim section As Section = doc.LastSection
			section.Body.AppendChild(para)
			'ExEnd
		End Sub

		<Test> _
		Public Sub RemoveSmartTagsFromCompositeNode()
			'ExStart
			'ExFor:CompositeNode.RemoveSmartTags
			'ExSummary:Removes all smart tags from descendant nodes of the composite node.
			Dim doc As New Document(MyDir & "Document.doc")

			' Remove smart tags from the first paragraph in the document.
			doc.FirstSection.Body.FirstParagraph.RemoveSmartTags()
			'ExEnd
		End Sub

		<Test> _
		Public Sub GetIndexOfNode()
			'ExStart
			'ExFor:CompositeNode.IndexOf
			'ExSummary:Shows how to get the index of a given child node from its parent.
			Dim doc As New Document(MyDir & "Rendering.doc")

			' Get the body of the first section in the document.
			Dim body As Body = doc.FirstSection.Body
			' Retrieve the index of the last paragraph in the body.
			Dim index As Integer = body.ChildNodes.IndexOf(body.LastParagraph)
			'ExEnd

			' Verify that the index is correct.
			Assert.AreEqual(24, index)
		End Sub

		<Test> _
		Public Sub GetNodeTypeEnums()
			'ExStart
			'ExFor:Paragraph.NodeType
			'ExFor:Table.NodeType
			'ExFor:Node.NodeType
			'ExFor:Footnote.NodeType
			'ExFor:FormField.NodeType
			'ExFor:SmartTag.NodeType
			'ExFor:Cell.NodeType
			'ExFor:Row.NodeType
			'ExFor:Document.NodeType
			'ExFor:Comment.NodeType
			'ExFor:Run.NodeType
			'ExFor:Section.NodeType
			'ExFor:SpecialChar.NodeType
			'ExFor:Shape.NodeType
			'ExFor:FieldEnd.NodeType
			'ExFor:FieldSeparator.NodeType
			'ExFor:FieldStart.NodeType
			'ExFor:BookmarkStart.NodeType
			'ExFor:CommentRangeEnd.NodeType
			'ExFor:BuildingBlock.NodeType
			'ExFor:GlossaryDocument.NodeType
			'ExFor:BookmarkEnd.NodeType
			'ExFor:GroupShape.NodeType
			'ExFor:CommentRangeStart.NodeType
			'ExId:GetNodeTypeEnums
			'ExSummary:Shows how to retrieve the NodeType enumeration of nodes.
			Dim doc As New Document(MyDir & "Document.doc")

			' Let's pick a node that we can't be quite sure of what type it is.
			' In this case lets pick the first node of the first paragraph in the body of the document
			Dim node As Node = doc.FirstSection.Body.FirstParagraph.FirstChild
			Console.WriteLine("NodeType of first child: " & Node.NodeTypeToString(node.NodeType))

			' This time let's pick a node that we know the type of. Create a new paragraph and a table node.
			Dim para As New Paragraph(doc)
			Dim table As New Table(doc)

			' Access to NodeType for typed nodes will always return their specific NodeType. 
			' i.e A paragraph node will always return NodeType.Paragraph, a table node will always return NodeType.Table.
			Console.WriteLine("NodeType of Paragraph: " & Node.NodeTypeToString(para.NodeType))
			Console.WriteLine("NodeType of Table: " & Node.NodeTypeToString(table.NodeType))
			'ExEnd
		End Sub

		<Test> _
		Public Sub ConvertNodeToHtmlWithDefaultOptions()
			'ExStart
			'ExFor:Node.ToString(SaveFormat)
			'ExSummary:Exports the content of a node to string in HTML format using default options.
			Dim doc As New Document(MyDir & "Document.doc")

			' Extract the last paragraph in the document to convert to HTML.
			Dim node As Node = doc.LastSection.Body.LastParagraph

			' When ToString is called using the SaveFormat overload then conversion is executed using default save options. 
			' When saving to HTML using default options the following settings are set:
			'   ExportImagesAsBase64 = true
			'   CssStyleSheetType = CssStyleSheetType.Inline
			'   ExportFontResources = false
			Dim nodeAsHtml As String = node.ToString(SaveFormat.Html)
			'ExEnd

			Assert.AreEqual("<p style=""margin-top:0pt; margin-bottom:0pt; font-size:12pt""><span style=""font-family:'Times New Roman'"">Hello World!</span></p>", nodeAsHtml)
		End Sub

		<Test> _
		Public Sub ConvertNodeToHtmlWithSaveOptions()
			'ExStart
			'ExFor:Node.ToString(SaveOptions)
			'ExSummary:Exports the content of a node to string in HTML format using custom specified options.
			Dim doc As New Document(MyDir & "Document.doc")

			' Extract the last paragraph in the document to convert to HTML.
			Dim node As Node = doc.LastSection.Body.LastParagraph

			' Create an instance of HtmlSaveOptions and set a few options.
			Dim saveOptions As New HtmlSaveOptions()
			saveOptions.ExportHeadersFootersMode = ExportHeadersFootersMode.PerSection
			saveOptions.ExportRelativeFontSize = True

			' Convert the document to HTML and return as a string. Pass the instance of HtmlSaveOptions to
			' to use the specified options during the conversion.
			Dim nodeAsHtml As String = node.ToString(saveOptions)
			'ExEnd

			Assert.AreEqual("<p style=""margin-top:0pt; margin-bottom:0pt""><span style=""font-family:'Times New Roman'"">Hello World!</span></p>", nodeAsHtml)
		End Sub

		<Test> _
		Public Sub TypedNodeCollectionToArray()
			Dim doc As New Document()

			'ExStart
			'ExFor:ParagraphCollection.ToArray
			'ExSummary:Demonstrates typed implementations of ToArray on classes derived from NodeCollection.
			' You can use ToArray to return a typed array of nodes.
			Dim paras() As Paragraph = doc.FirstSection.Body.Paragraphs.ToArray()
			'ExEnd

			Assert.Greater(paras.Length, 0)
		End Sub

		<Test> _
		Public Sub NodeEnumerationHotRemove()
			'ExStart
			'ExFor:ParagraphCollection.ToArray
			'ExSummary:Demonstrates how to use "hot remove" to remove a node during enumeration.
			Dim builder As New DocumentBuilder()
			builder.Writeln("The first paragraph")
			builder.Writeln("The second paragraph")
			builder.Writeln("The third paragraph")
			builder.Writeln("The fourth paragraph")

			' Hot remove allows a node to be removed from a live collection and have the enumeration continue.
			For Each para As Paragraph In builder.Document.FirstSection.Body.GetChildNodes(NodeType.Paragraph, True)
				If para.Range.Text.Contains("third") Then
					' Enumeration will continue even after this node is removed.
					para.Remove()
				End If
			Next para
			'ExEnd
		End Sub

		<Test> _
		Public Sub EnumerationHotRemoveLimitations()
			'ExStart
			'ExFor:ParagraphCollection.ToArray
			'ExSummary:Demonstrates an example breakage of the node collection enumerator.
			Dim builder As New DocumentBuilder()
			builder.Writeln("The first paragraph")
			builder.Writeln("The second paragraph")
			builder.Writeln("The third paragraph")
			builder.Writeln("The fourth paragraph")

			' This causes unexpected behavior, the fourth pargraph in the collection is not visited.
			For Each para As Paragraph In builder.Document.FirstSection.Body.GetChildNodes(NodeType.Paragraph, True)
				If para.Range.Text.Contains("third") Then
					para.PreviousSibling.Remove()
					para.Remove()
				End If
			Next para
			'ExEnd
		End Sub
	End Class
End Namespace
