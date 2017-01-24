' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports System.Text.RegularExpressions

Imports Aspose.Words

Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Public Class ExRange
		Inherits ApiExampleBase
		<Test> _
		Public Sub DeleteSelection()
			'ExStart
			'ExFor:Node.Range
			'ExFor:Range.Delete
			'ExSummary:Shows how to delete a section from a Word document.
			' Open Word document.
			Dim doc As New Document(MyDir & "Range.DeleteSection.doc")

			' The document contains two sections. Each section has a paragraph of text.
			Console.WriteLine(doc.GetText())

			' Delete the first section from the document.
			doc.Sections(0).Range.Delete()

			' Check the first section was deleted by looking at the text of the whole document again.
			Console.WriteLine(doc.GetText())
			'ExEnd

			Assert.AreEqual("Hello2" & Constants.vbFormFeed, doc.GetText())
		End Sub

		<Test> _
		Public Sub ReplaceSimple()
			'ExStart
			'ExFor:Range.Replace(String,String,Boolean,Boolean)
			'ExSummary:Simple find and replace operation.
			' Open the document.
			Dim doc As New Document(MyDir & "Range.ReplaceSimple.doc")

			' Check the document contains what we are about to test.
			Console.WriteLine(doc.FirstSection.Body.Paragraphs(0).GetText())

			' Replace the text in the document.
			doc.Range.Replace("_CustomerName_", "James Bond", False, False)

			' Save the modified document.
			doc.Save(MyDir & "\Artifacts\Range.ReplaceSimple.doc")
			'ExEnd

			Assert.AreEqual("Hello James Bond," & Constants.vbCr + Constants.vbFormFeed, doc.GetText())
		End Sub

		''' <summary>
		''' This calls the below method to resolve skipping of [Test] in VB.NET.
		''' </summary>
		<Test> _
		Public Sub ReplaceWithInsertHtmlCaller()
			Me.ReplaceWithInsertHtml()
		End Sub

		'ExStart
		'ExFor:Range.Replace(Regex,IReplacingCallback,Boolean)
		'ExFor:ReplacingArgs.Replacement
		'ExFor:IReplacingCallback
		'ExFor:IReplacingCallback.Replacing
		'ExFor:ReplacingArgs
		'ExFor:DocumentBuilder.InsertHtml(string)
		'ExSummary:Replaces text specified with regular expression with HTML.
		Public Sub ReplaceWithInsertHtml()
			' Open the document.
			Dim doc As New Document(MyDir & "Range.ReplaceWithInsertHtml.doc")

			doc.Range.Replace(New Regex("<CustomerName>"), New ReplaceWithHtmlEvaluator(), False)

			' Save the modified document.
			doc.Save(MyDir & "\Artifacts\Range.ReplaceWithInsertHtml.doc")

			Assert.AreEqual("Hello James Bond," & Constants.vbCr + Constants.vbFormFeed, doc.GetText()) 'ExSkip
		End Sub

		Private Class ReplaceWithHtmlEvaluator
			Implements IReplacingCallback
			''' <summary>
			''' NOTE: This is a simplistic method that will only work well when the match
			''' starts at the beginning of a run.
			''' </summary>
			Private Function IReplacingCallback_Replacing(ByVal e As ReplacingArgs) As ReplaceAction Implements IReplacingCallback.Replacing
				Dim builder As New DocumentBuilder(CType(e.MatchNode.Document, Document))
				builder.MoveTo(e.MatchNode)
				' Replace '<CustomerName>' text with a red bold name.
				builder.InsertHtml("<b><font color='red'>James Bond</font></b>")

				e.Replacement = ""
				Return ReplaceAction.Replace
			End Function
		End Class
		'ExEnd

		<Test> _
		Public Sub RangesGetText()
			'ExStart
			'ExFor:Range
			'ExFor:Range.Text
			'ExId:RangesGetText
			'ExSummary:Shows how to get plain, unformatted text of a range.
			Dim doc As New Document(MyDir & "Document.doc")
			Dim text As String = doc.Range.Text
			'ExEnd
		End Sub

		<Test> _
		Public Sub ReplaceWithString()
			'ExStart
			'ExFor:Range
			'ExId:RangesReplaceString
			'ExSummary:Shows how to replace all occurrences of word "sad" to "bad".
			Dim doc As New Document(MyDir & "Document.doc")
			doc.Range.Replace("sad", "bad", False, True)
			'ExEnd
			doc.Save(MyDir & "\Artifacts\ReplaceWithString.doc")
		End Sub

		<Test> _
		Public Sub ReplaceWithRegex()
			'ExStart
			'ExFor:Range.Replace(Regex, String)
			'ExId:RangesReplaceRegex
			'ExSummary:Shows how to replace all occurrences of words "sad" or "mad" to "bad".
			Dim doc As New Document(MyDir & "Document.doc")
			doc.Range.Replace(New Regex("[s|m]ad"), "bad")
			'ExEnd
			doc.Save(MyDir & "\Artifacts\ReplaceWithRegex.doc")
		End Sub

		''' <summary>
		''' This calls the below method to resolve skipping of [Test] in VB.NET.
		''' </summary>
		<Test> _
		Public Sub ReplaceWithEvaluatorCaller()
			Me.ReplaceWithEvaluator()
		End Sub

		'ExStart
		'ExFor:Range
		'ExFor:ReplacingArgs.Match
		'ExId:RangesReplaceWithReplaceEvaluator
		'ExSummary:Shows how to replace with a custom evaluator.
		Public Sub ReplaceWithEvaluator()
			Dim doc As New Document(MyDir & "Range.ReplaceWithEvaluator.doc")
			doc.Range.Replace(New Regex("[s|m]ad"), New MyReplaceEvaluator(), True)
			doc.Save(MyDir & "\Artifacts\Range.ReplaceWithEvaluator.doc")
		End Sub

		Private Class MyReplaceEvaluator
			Implements IReplacingCallback
			''' <summary>
			''' This is called during a replace operation each time a match is found.
			''' This method appends a number to the match string and returns it as a replacement string.
			''' </summary>
			Private Function IReplacingCallback_Replacing(ByVal e As ReplacingArgs) As ReplaceAction Implements IReplacingCallback.Replacing
				e.Replacement = e.Match.ToString() & Me.mMatchNumber.ToString()
				Me.mMatchNumber += 1
				Return ReplaceAction.Replace
			End Function

			Private mMatchNumber As Integer
		End Class
		'ExEnd

		<Test> _
		Public Sub RangesDeleteText()
			'ExStart
			'ExId:RangesDeleteText
			'ExSummary:Shows how to delete all characters of a range.
			Dim doc As New Document(MyDir & "Document.doc")
			doc.Sections(0).Range.Delete()
			'ExEnd
		End Sub

		''' <summary>
		''' RK This works, but the logic is so complicated that I don't want to show it to users.
		''' </summary>
		<Test> _
		Public Sub ChangeTextToHyperlinks()
			Dim doc As New Document(MyDir & "Range.ChangeTextToHyperlinks.doc")

			' Create regular expression for URL search
			Dim regexUrl As New Regex("(?<Protocol>\w+):\/\/(?<Domain>[\w.]+\/?)\S*(?x)")

			' Run replacement, using regular expression and evaluator.
			doc.Range.Replace(regexUrl, New ChangeTextToHyperlinksEvaluator(doc), False)

			' Save updated document.
			doc.Save(MyDir & "\Artifacts\Range.ChangeTextToHyperlinks.docx")
		End Sub

		Private Class ChangeTextToHyperlinksEvaluator
			Implements IReplacingCallback
			Friend Sub New(ByVal doc As Document)
				Me.mBuilder = New DocumentBuilder(doc)
			End Sub

			Private Function IReplacingCallback_Replacing(ByVal e As ReplacingArgs) As ReplaceAction Implements IReplacingCallback.Replacing
				' This is the run node that contains the found text. Note that the run might contain other 
				' text apart from the URL. All the complexity below is just to handle that. I don't think there
				' is a simpler way at the moment.
				Dim run As Run = CType(e.MatchNode, Run)

				Dim para As Paragraph = run.ParentParagraph

				Dim url As String = e.Match.Value

				' We are using \xbf (inverted question mark) symbol for temporary purposes.
				' Any symbol will do that is non-special and is guaranteed not to be presented in the document.
				' The purpose is to split the matched run into two and insert a hyperlink field between them.
				para.Range.Replace(url, ChrW(&Hbf).ToString(), True, True)

				Dim subRun As Run = CType(run.Clone(False), Run)
				Dim pos As Integer = run.Text.IndexOf(ChrW(&Hbf).ToString())
				subRun.Text = subRun.Text.Substring(0, pos)
				run.Text = run.Text.Substring(pos + 1, run.Text.Length - pos - 1)

				para.ChildNodes.Insert(para.ChildNodes.IndexOf(run), subRun)

				Me.mBuilder.MoveTo(run)

				' Specify font formatting for the hyperlink.
				Me.mBuilder.Font.Color = Color.Blue
				Me.mBuilder.Font.Underline = Underline.Single

				' Insert the hyperlink.
				Me.mBuilder.InsertHyperlink(url, url, False)

				' Clear hyperlink formatting.
				Me.mBuilder.Font.ClearFormatting()

				' Let's remove run if it is empty.
				If run.Text.Equals("") Then
					run.Remove()
				End If

				' No replace action is necessary - we have already done what we intended to do.
				Return ReplaceAction.Skip
			End Function

			Private ReadOnly mBuilder As DocumentBuilder
		End Class
	End Class
End Namespace
