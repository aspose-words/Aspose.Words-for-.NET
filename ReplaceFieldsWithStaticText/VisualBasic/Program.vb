'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Reflection
Imports System.Collections

Imports Aspose.Words
Imports Aspose.Words.Fields
Imports Aspose.Words.Tables
Imports System.Diagnostics

Namespace ReplaceFieldsWithStaticText
	Friend Class Program
		Public Shared Sub Main(ByVal args() As String)
			' Sample infrastructure.
			Dim exeDir As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar
			Dim dataDir As String = New Uri(New Uri(exeDir), "../../Data/").LocalPath

			' Call a method to show how to convert all IF fields in a document to static text.
			ConvertFieldsInDocument(dataDir)
			' Reload the document and this time convert all PAGE fields only encountered in the first body of the document.
			ConvertFieldsInBody(dataDir)
			' Reload the document again and convert only the IF field in the last paragraph to static text.
			ConvertFieldsInParagraph(dataDir)
		End Sub

		'ExStart:
		'ExFor:DocumentVisitor.VisitTableStart(Aspose.Words.Tables.Table)
		'ExId:ConvertFieldsToStaticText
		'ExSummary:This class provides a static method convert fields of a particular type to static text.
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
					Debug.Assert(Not(TypeOf node Is FieldChar AndAlso (CType(node, FieldChar)).FieldType.Equals(targetFieldType)), "Error: A field node that should be removed still remains.") 'ExSkip
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
		'ExEnd

		Public Shared Sub ConvertFieldsInDocument(ByVal dataDir As String)
			'ExStart:
			'ExId:FieldsToStaticTextDocument
			'ExSummary:Shows how to convert all fields of a specified type in a document to static text.
			Dim doc As New Document(dataDir & "TestFile.doc")

			' Pass the appropriate parameters to convert all IF fields encountered in the document (including headers and footers) to static text.
			FieldsHelper.ConvertFieldsToStaticText(doc, FieldType.FieldIf)

			' Save the document with fields transformed to disk.
			doc.Save(dataDir & "TestFileDocument Out.doc")
			'ExEnd
		End Sub

		Public Shared Sub ConvertFieldsInBody(ByVal dataDir As String)
			'ExStart:
			'ExId:FieldsToStaticTextBody
			'ExSummary:Shows how to convert all fields of a specified type in a body of a document to static text.
			Dim doc As New Document(dataDir & "TestFile.doc")

			' Pass the appropriate parameters to convert PAGE fields encountered to static text only in the body of the first section.
			FieldsHelper.ConvertFieldsToStaticText(doc.FirstSection.Body, FieldType.FieldPage)

			' Save the document with fields transformed to disk.
			doc.Save(dataDir & "TestFileBody Out.doc")
			'ExEnd
		End Sub

		Public Shared Sub ConvertFieldsInParagraph(ByVal dataDir As String)
			'ExStart:
			'ExId:FieldsToStaticTextParagraph
			'ExSummary:Shows how to convert all fields of a specified type in a paragraph to static text.
			Dim doc As New Document(dataDir & "TestFile.doc")

			' Pass the appropriate parameters to convert all IF fields to static text that are encountered only in the last 
			' paragraph of the document.
			FieldsHelper.ConvertFieldsToStaticText(doc.FirstSection.Body.LastParagraph, FieldType.FieldIf)

			' Save the document with fields transformed to disk.
			doc.Save(dataDir & "TestFileParagraph Out.doc")
			'ExEnd
		End Sub

	End Class
End Namespace
