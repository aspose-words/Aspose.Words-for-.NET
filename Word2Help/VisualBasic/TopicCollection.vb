'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////
'14/9/06 by Vladimir Averkin

Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Collections
Imports System.Text
Imports System.Xml
Imports Aspose.Words

Namespace Word2Help
	''' <summary>
	''' This is the main class.
	''' Loads Word document(s), splits them into topics, saves HTML files and builds content.xml.
	''' </summary>
	Public Class TopicCollection
		''' <summary>
		''' Ctor.
		''' </summary>
		''' <param name="htmlTemplatesDir">The directory that contains header.html, banner.html and footer.html files.</param>
		''' <param name="fixUrl">The url that will be removed from any hyperlinks that start with this url.
		''' This allows turning some absolute URLS into relative.</param>
		Public Sub New(ByVal htmlTemplatesDir As String, ByVal fixUrl As String)
			mTopics = New ArrayList()
			mFixUrl = fixUrl
			mHtmlHeader = ReadFile(Path.Combine(htmlTemplatesDir, "header.html"))
			mHtmlBanner = ReadFile(Path.Combine(htmlTemplatesDir, "banner.html"))
			mHtmlFooter = ReadFile(Path.Combine(htmlTemplatesDir, "footer.html"))
		End Sub

		''' <summary>
		''' Processes all DOC files found in the specified directory.
		''' Loads and splits each document into separate topics.
		''' </summary>
		Public Sub AddFromDir(ByVal dirName As String)
			For Each filename As String In Directory.GetFiles(dirName, "*.doc")
				AddFromFile(filename)
			Next filename
		End Sub

		''' <summary>
		''' Processes a specified DOC file. Loads and splits into topics.
		''' </summary>
		Public Sub AddFromFile(ByVal fileName As String)
			Dim doc As New Document(fileName)
			InsertTopicSections(doc)
			AddTopics(doc)
		End Sub

		''' <summary>
		''' Saves all topics as HTML files.
		''' </summary>
		Public Sub WriteHtml(ByVal outDir As String)
			For Each topic As Topic In mTopics
				If (Not topic.IsHeadingOnly) Then
					topic.WriteHtml(mHtmlHeader, mHtmlBanner, mHtmlFooter, outDir)
				End If
			Next topic
		End Sub

		''' <summary>
		''' Saves the content.xml file that describes the tree of topics.
		''' </summary>
		Public Sub WriteContentXml(ByVal outDir As String)
			Dim writer As New XmlTextWriter(Path.Combine(outDir, "content.xml"), Encoding.UTF8)
			writer.Namespaces = False
			writer.Formatting = Formatting.Indented

			writer.WriteStartDocument(True)
			writer.WriteStartElement("content")
			writer.WriteAttributeString("dir", outDir)

			For i As Integer = 0 To mTopics.Count - 1
				Dim topic As Topic = CType(mTopics(i), Topic)

				Dim nextTopicIdx As Integer = i + 1
				Dim nextTopic As Topic = If((nextTopicIdx < mTopics.Count), CType(mTopics(i + 1), Topic), Nothing)

				Dim nextHeadingLevel As Integer = If((nextTopic IsNot Nothing), nextTopic.HeadingLevel, 0)

				If nextHeadingLevel > topic.HeadingLevel Then
					' Next topic is nested, therefore we have to start a book. 
					' We only allow increase level at a time.
					If nextHeadingLevel <> topic.HeadingLevel + 1 Then
						Throw New Exception("Topic is nested for more than one level at a time. Title: " & topic.Title)
					End If

					WriteBookStart(writer, topic)
				ElseIf nextHeadingLevel < topic.HeadingLevel Then
					' Next topic is one or more levels higher in the outline.

					' Write out the current topic.
					WriteItem(writer, topic.Title, topic.FileName)

					' End one or more nested topics could have ended at this point.
					Dim levelsToClose As Integer = topic.HeadingLevel - nextHeadingLevel
					Do While levelsToClose > 0
						WriteBookEnd(writer)
						levelsToClose -= 1
					Loop
				Else
					' A topic at the current level and it has no children.
					WriteItem(writer, topic.Title, topic.FileName)
				End If
			Next i

			writer.WriteEndElement() ' content
			writer.WriteEndDocument()
			writer.Flush()
			writer.Close()
		End Sub

		''' <summary>
		''' Inserts section breaks that delimit the topics.
		''' </summary>
		''' <param name="doc">The document where to insert the section breaks.</param>
		Private Shared Sub InsertTopicSections(ByVal doc As Document)
			Dim builder As New DocumentBuilder(doc)

			Dim paras As NodeCollection = doc.GetChildNodes(NodeType.Paragraph, True, False)
			Dim topicStartParas As New ArrayList()

			For Each para As Paragraph In paras
				Dim style As StyleIdentifier = para.ParagraphFormat.StyleIdentifier
				If (style >= StyleIdentifier.Heading1) AndAlso (style <= MaxTopicHeading) AndAlso (para.HasChildNodes) Then
					' Select heading paragraphs that must become topic starts.
					' We can't modify them in this loop, we have to remember them in an array first.
					topicStartParas.Add(para)
				ElseIf (style > MaxTopicHeading) AndAlso (style <= StyleIdentifier.Heading9) Then
					' Pull up headings. For example: if Heading 1-4 become topics, then I want Headings 5+ 
					' to become Headings 4+. Maybe I want to pull up even higher?
					para.ParagraphFormat.StyleIdentifier = CType(CInt(Fix(style)) - 1, StyleIdentifier)
				End If
			Next para

			For Each para As Paragraph In topicStartParas
				Dim section As Section = para.ParentSection

				' Insert section break if the paragraph is not at the beginning of a section already.
				If para IsNot section.Body.FirstParagraph Then
					builder.MoveTo(para.FirstChild)
					builder.InsertBreak(BreakType.SectionBreakNewPage)

					' This is the paragraph that was inserted at the end of the now old section.
					' We don't really need the extra paragraph, we just needed the section.
					section.Body.LastParagraph.Remove()
				End If
			Next para
		End Sub

		''' <summary>
		''' Goes through the sections in the document and adds them as topics to the collection.
		''' </summary>
		Private Sub AddTopics(ByVal doc As Document)
			For Each section As Section In doc.Sections
				Try
					Dim topic As New Topic(section, mFixUrl)
					mTopics.Add(topic)
				Catch e As Exception
					' If one topic fails, we continue with others.
					Console.WriteLine(e.Message)
				End Try
			Next section
		End Sub

		Private Shared Sub WriteBookStart(ByVal writer As XmlWriter, ByVal topic As Topic)
			writer.WriteStartElement("book")
			writer.WriteAttributeString("name", topic.Title)

			If (Not topic.IsHeadingOnly) Then
				writer.WriteAttributeString("href", topic.FileName)
			End If
		End Sub

		Private Shared Sub WriteBookEnd(ByVal writer As XmlWriter)
			writer.WriteEndElement() ' book
		End Sub

		Private Shared Sub WriteItem(ByVal writer As XmlWriter, ByVal name As String, ByVal href As String)
			writer.WriteStartElement("item")
			writer.WriteAttributeString("name", name)
			writer.WriteAttributeString("href", href)
			writer.WriteEndElement() ' item
		End Sub

		Private Shared Function ReadFile(ByVal fileName As String) As String
			Using reader As New StreamReader(fileName)
				Return reader.ReadToEnd()
			End Using
		End Function

		Private ReadOnly mTopics As ArrayList
		Private ReadOnly mFixUrl As String
		Private ReadOnly mHtmlHeader As String
		Private ReadOnly mHtmlBanner As String
		Private ReadOnly mHtmlFooter As String

		''' <summary>
		''' Specifies the maximum Heading X number. 
		''' All of the headings above or equal to this will be put into their own topics.
		''' </summary>
		Private Const MaxTopicHeading As StyleIdentifier = StyleIdentifier.Heading4
	End Class
End Namespace
