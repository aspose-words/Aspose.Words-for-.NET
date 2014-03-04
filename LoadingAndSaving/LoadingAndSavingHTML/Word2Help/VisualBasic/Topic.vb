'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Text
Imports Aspose.Words
Imports Aspose.Words.Fields
Imports Aspose.Words.Saving

Namespace Word2HelpExample
	''' <summary>
	''' Represents a single topic that will be written as an HTML file.
	''' </summary>
	Public Class Topic
		''' <summary>
		''' Creates a topic.
		''' </summary>
		Public Sub New(ByVal section As Section, ByVal fixUrl As String)
			mTopicDoc = New Document()
			mTopicDoc.AppendChild(mTopicDoc.ImportNode(section, True, ImportFormatMode.KeepSourceFormatting))
			mTopicDoc.FirstSection.Remove()

			Dim headingPara As Paragraph = CType(mTopicDoc.FirstSection.Body.FirstChild, Paragraph)
			If headingPara Is Nothing Then
				ThrowTopicException("The section does not start with a paragraph.", section)
			End If

			mHeadingLevel = headingPara.ParagraphFormat.StyleIdentifier - StyleIdentifier.Heading1
			If (mHeadingLevel < 0) OrElse (mHeadingLevel > 8) Then
				ThrowTopicException("This topic does not start with a heading style paragraph.", section)
			End If

			mTitle = headingPara.GetText().Trim()
			If mTitle = "" Then
				ThrowTopicException("This topic heading does not have text.", section)
			End If

			' We actually remove the heading paragraph because <h1> will be output in the banner.
			headingPara.Remove()

			mTopicDoc.BuiltInDocumentProperties.Title = mTitle

			FixHyperlinks(section.Document, fixUrl)
		End Sub

		Private Shared Sub ThrowTopicException(ByVal message As String, ByVal section As Section)
			Throw New Exception(message & " Section text: " & section.Body.ToString(SaveFormat.Text).Substring(0, 50))
		End Sub

		Private Sub FixHyperlinks(ByVal originalDoc As DocumentBase, ByVal fixUrl As String)
			If fixUrl.EndsWith("/") Then
				fixUrl = fixUrl.Substring(0, fixUrl.Length - 1)
			End If

			Dim fieldStarts As NodeCollection = mTopicDoc.GetChildNodes(NodeType.FieldStart, True)
			For Each fieldStart As FieldStart In fieldStarts
				If fieldStart.FieldType <> FieldType.FieldHyperlink Then
					Continue For
				End If

				Dim hyperlink As New Hyperlink(fieldStart)
				If hyperlink.IsLocal Then
					' We use "Hyperlink to a place in this document" feature of Microsoft Word
					' to create local hyperlinks between topics within the same doc file.
					' It causes MS Word to auto generate the bookmark name.
					Dim bmkName As String = hyperlink.Target

					' But we have to follow the bookmark to get the text of the topic heading paragraph
					' in order to be able to build the proper filename of the topic file.
					Dim bmk As Bookmark = originalDoc.Range.Bookmarks(bmkName)

					If bmk Is Nothing Then
						Throw New Exception(String.Format("Found a link to a bookmark, but cannot locate the bookmark. Name:'{0}'.", bmkName))
					End If

					Dim para As Paragraph = CType(bmk.BookmarkStart.ParentNode, Paragraph)
					Dim topicName As String = para.GetText().Trim()

					hyperlink.Target = HeadingToFileName(topicName) & ".html"
					hyperlink.IsLocal = False
				Else
					' We "fix" URL like this:
					' http://www.aspose.com/Products/Aspose.Words/Api/Aspose.Words.Body.html
					' by changing them into this:
					' Aspose.Words.Body.html
					If hyperlink.Target.StartsWith(fixUrl) AndAlso (hyperlink.Target.Length > (fixUrl.Length + 1)) Then
						hyperlink.Target = hyperlink.Target.Substring(fixUrl.Length + 1)
					End If
				End If
			Next fieldStart
		End Sub

		Public Sub WriteHtml(ByVal htmlHeader As String, ByVal htmlBanner As String, ByVal htmlFooter As String, ByVal outDir As String)
			Dim fileName As String = Path.Combine(outDir, Me.FileName)

			Dim saveOptions As New HtmlSaveOptions()
			saveOptions.PrettyFormat = True
			' This is to allow headings to appear to the left of main text.
			saveOptions.AllowNegativeLeftIndent = True
			' Disable headers and footers.
			saveOptions.ExportHeadersFootersMode = ExportHeadersFootersMode.None

			' Export the document to HTML.
			mTopicDoc.Save(fileName, saveOptions)

			' We need to modify the HTML string, read HTML back.
			Dim html As String
			Using reader As New StreamReader(fileName)
				html = reader.ReadToEnd()
			End Using

			' Builds the HTML <head> element.
			Dim header As String = RegularExpressions.HtmlTitle.Replace(htmlHeader, mTitle, 1)

			' Applies the new <head> element instead of the original one.
			html = RegularExpressions.HtmlHead.Replace(html, header, 1)
			html = RegularExpressions.HtmlBodyDivStart.Replace(html, " id=""nstext""", 1)

			Dim banner As String = htmlBanner.Replace("###TOPIC_NAME###", mTitle)

			' Add the standard banner.
			html = html.Replace("<body>", "<body>" & banner)

			' Add the standard footer.
			html = html.Replace("</body>", htmlFooter & "</body>")

			Using writer As New StreamWriter(fileName)
				writer.Write(html)
			End Using
		End Sub

		''' <summary>
		''' Removes various characters from the heading to form a file name that does not require escaping.
		''' </summary>
		Private Shared Function HeadingToFileName(ByVal heading As String) As String
			Dim b As New StringBuilder()
			For Each c As Char In heading
				If Char.IsLetterOrDigit(c) Then
					b.Append(c)
				End If
			Next c

			Return b.ToString()
		End Function

		Public ReadOnly Property Document() As Document
			Get
				Return mTopicDoc
			End Get
		End Property

		''' <summary>
		''' Gets the name of the topic html file without path.
		''' </summary>
		Public ReadOnly Property FileName() As String
			Get
				Return HeadingToFileName(mTitle) & ".html"
			End Get
		End Property

		Public ReadOnly Property Title() As String
			Get
				Return mTitle
			End Get
		End Property

		Public ReadOnly Property HeadingLevel() As Integer
			Get
				Return mHeadingLevel
			End Get
		End Property

		''' <summary>
		''' Returns true if the topic has no text (the heading paragraph has already been removed from the topic).
		''' </summary>
		Public ReadOnly Property IsHeadingOnly() As Boolean
			Get
				Dim body As Body = mTopicDoc.FirstSection.Body
				Return (body.FirstParagraph Is Nothing)
			End Get
		End Property

		Private ReadOnly mTopicDoc As Document
		Private ReadOnly mTitle As String
		Private ReadOnly mHeadingLevel As Integer
	End Class
End Namespace