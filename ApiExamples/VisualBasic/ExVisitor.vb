' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports System.Text

Imports Aspose.Words
Imports Aspose.Words.Fields

Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Public Class ExVisitor
		Inherits ApiExampleBase
		''' <summary>
		''' This calls the below method to resolve skipping of [Test] in VB.NET.
		''' </summary>
		<Test> _
		Public Sub ToTextCaller()
			Me.ToText()
		End Sub

		'ExStart
		'ExFor:Document.Accept
		'ExFor:Body.Accept
		'ExFor:DocumentVisitor
		'ExFor:DocumentVisitor.VisitAbsolutePositionTab
		'ExFor:DocumentVisitor.VisitBookmarkStart 
		'ExFor:DocumentVisitor.VisitBookmarkEnd
		'ExFor:DocumentVisitor.VisitRun
		'ExFor:DocumentVisitor.VisitFieldStart
		'ExFor:DocumentVisitor.VisitFieldEnd
		'ExFor:DocumentVisitor.VisitFieldSeparator
		'ExFor:DocumentVisitor.VisitBodyStart
		'ExFor:DocumentVisitor.VisitBodyEnd
		'ExFor:DocumentVisitor.VisitParagraphEnd
		'ExFor:DocumentVisitor.VisitHeaderFooterStart
		'ExFor:VisitorAction
		'ExId:ExtractContentDocToTxtConverter
		'ExSummary:Shows how to use the Visitor pattern to add new operations to the Aspose.Words object model. In this case we create a simple document converter into a text format.
		Public Sub ToText()
			' Open the document we want to convert.
			Dim doc As New Document(MyDir & "Visitor.ToText.doc")

			' Create an object that inherits from the DocumentVisitor class.
			Dim myConverter As New MyDocToTxtWriter()

			' This is the well known Visitor pattern. Get the model to accept a visitor.
			' The model will iterate through itself by calling the corresponding methods
			' on the visitor object (this is called visiting).
			' 
			' Note that every node in the object model has the Accept method so the visiting
			' can be executed not only for the whole document, but for any node in the document.
			doc.Accept(myConverter)

			' Once the visiting is complete, we can retrieve the result of the operation,
			' that in this example, has accumulated in the visitor.
			Console.WriteLine(myConverter.GetText())
		End Sub

		''' <summary>
		''' Simple implementation of saving a document in the plain text format. Implemented as a Visitor.
		''' </summary>
		Public Class MyDocToTxtWriter
			Inherits DocumentVisitor
			Public Sub New()
				Me.mIsSkipText = False
				Me.mBuilder = New StringBuilder()
			End Sub

			''' <summary>
			''' Gets the plain text of the document that was accumulated by the visitor.
			''' </summary>
			Public Function GetText() As String
				Return Me.mBuilder.ToString()
			End Function

			''' <summary>
			''' Called when a Run node is encountered in the document.
			''' </summary>
			Public Overrides Function VisitRun(ByVal run As Run) As VisitorAction
				Me.AppendText(run.Text)

				' Let the visitor continue visiting other nodes.
				Return VisitorAction.Continue
			End Function

			''' <summary>
			''' Called when a FieldStart node is encountered in the document.
			''' </summary>
			Public Overrides Function VisitFieldStart(ByVal fieldStart As FieldStart) As VisitorAction
				' In Microsoft Word, a field code (such as "MERGEFIELD FieldName") follows
				' after a field start character. We want to skip field codes and output field 
				' result only, therefore we use a flag to suspend the output while inside a field code.
				'
				' Note this is a very simplistic implementation and will not work very well
				' if you have nested fields in a document. 
				Me.mIsSkipText = True

				Return VisitorAction.Continue
			End Function

			''' <summary>
			''' Called when a FieldSeparator node is encountered in the document.
			''' </summary>
			Public Overrides Function VisitFieldSeparator(ByVal fieldSeparator As FieldSeparator) As VisitorAction
				' Once reached a field separator node, we enable the output because we are
				' now entering the field result nodes.
				Me.mIsSkipText = False

				Return VisitorAction.Continue
			End Function

			''' <summary>
			''' Called when a FieldEnd node is encountered in the document.
			''' </summary>
			Public Overrides Function VisitFieldEnd(ByVal fieldEnd As FieldEnd) As VisitorAction
				' Make sure we enable the output when reached a field end because some fields
				' do not have field separator and do not have field result.
				Me.mIsSkipText = False

				Return VisitorAction.Continue
			End Function

			''' <summary>
			''' Called when visiting of a Paragraph node is ended in the document.
			''' </summary>
			Public Overrides Function VisitParagraphEnd(ByVal paragraph As Paragraph) As VisitorAction
				' When outputting to plain text we output Cr+Lf characters.
				Me.AppendText(ControlChar.CrLf)

				Return VisitorAction.Continue
			End Function

			Public Overrides Function VisitBodyStart(ByVal body As Body) As VisitorAction
				' We can detect beginning and end of all composite nodes such as Section, Body, 
				' Table, Paragraph etc and provide custom handling for them.
				Me.mBuilder.Append("*** Body Started ***" & Constants.vbCrLf)

				Return VisitorAction.Continue
			End Function

			Public Overrides Function VisitBodyEnd(ByVal body As Body) As VisitorAction
				Me.mBuilder.Append("*** Body Ended ***" & Constants.vbCrLf)
				Return VisitorAction.Continue
			End Function

			''' <summary>
			''' Called when a HeaderFooter node is encountered in the document.
			''' </summary>
			Public Overrides Function VisitHeaderFooterStart(ByVal headerFooter As HeaderFooter) As VisitorAction
				' Returning this value from a visitor method causes visiting of this
				' node to stop and move on to visiting the next sibling node.
				' The net effect in this example is that the text of headers and footers
				' is not included in the resulting output.
				Return VisitorAction.SkipThisNode
			End Function

			''' <summary>
			''' Called when an AbsolutePositionTab is encountered in the document.
			''' </summary>
			Public Overrides Function VisitAbsolutePositionTab(ByVal tab As AbsolutePositionTab) As VisitorAction
				Me.mBuilder.Append(Constants.vbTab)
				Return VisitorAction.Continue
			End Function

			''' <summary>
			''' Called when a BookmarkStart is encountered in the document.
			''' </summary>
			Public Overrides Function VisitBookmarkStart(ByVal bookmarkStart As BookmarkStart) As VisitorAction
				Me.mBuilder.Append("[")
				Return VisitorAction.Continue
			End Function

			''' <summary>
			''' Called when a BookmarkEnd is encountered in the document.
			''' </summary>
			Public Overrides Function VisitBookmarkEnd(ByVal bookmarkEnd As BookmarkEnd) As VisitorAction
				Me.mBuilder.Append("]")
				Return VisitorAction.Continue
			End Function

			''' <summary>
			''' Adds text to the current output. Honours the enabled/disabled output flag.
			''' </summary>
			Private Sub AppendText(ByVal text As String)
				If (Not Me.mIsSkipText) Then
					Me.mBuilder.Append(text)
				End If
			End Sub

			Private ReadOnly mBuilder As StringBuilder
			Private mIsSkipText As Boolean
		End Class
		'ExEnd
	End Class
End Namespace