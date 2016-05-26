' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System

Imports Aspose.Words

Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Public Class ExInlineStory
		Inherits ApiExampleBase
		<Test> _
		Public Sub AddFootnote()
			'ExStart
			'ExFor:Footnote
			'ExFor:InlineStory
			'ExFor:InlineStory.Paragraphs
			'ExFor:InlineStory.FirstParagraph
			'ExFor:FootnoteType
			'ExFor:Footnote.#ctor
			'ExSummary:Shows how to add a footnote to a paragraph in the document.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)
			builder.Write("Some text is added.")

			Dim footnote As New Footnote(doc, FootnoteType.Footnote)
			builder.CurrentParagraph.AppendChild(footnote)
			footnote.Paragraphs.Add(New Paragraph(doc))
			footnote.FirstParagraph.Runs.Add(New Run(doc, "Footnote text."))
			'ExEnd

			Assert.AreEqual("Footnote text.", doc.GetChildNodes(NodeType.Footnote, True)(0).ToString(SaveFormat.Text).Trim())
		End Sub

		<Test> _
		Public Sub AddComment()
			'ExStart
			'ExFor:Comment
			'ExFor:InlineStory
			'ExFor:InlineStory.Paragraphs
			'ExFor:InlineStory.FirstParagraph
			'ExFor:Comment.#ctor(DocumentBase, String, String, DateTime)
			'ExSummary:Shows how to add a comment to a paragraph in the document.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)
			builder.Write("Some text is added.")

			Dim comment As New Comment(doc, "Amy Lee", "AL", DateTime.Today)
			builder.CurrentParagraph.AppendChild(comment)
			comment.Paragraphs.Add(New Paragraph(doc))
			comment.FirstParagraph.Runs.Add(New Run(doc, "Comment text."))
			'ExEnd

			Assert.AreEqual("Comment text." & Constants.vbCr, (doc.GetChildNodes(NodeType.Comment, True)(0)).GetText())
		End Sub
	End Class

End Namespace
