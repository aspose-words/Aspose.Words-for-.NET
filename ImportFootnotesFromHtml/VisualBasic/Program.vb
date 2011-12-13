'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////
' 3/1/08 by Roman Korchagin

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports System.Text.RegularExpressions

Imports Aspose.Words
Imports Aspose.Words.Fields
Imports Aspose.Words.Saving


Namespace ImportFootnotesFromHtml
	''' <summary>
	''' This is a sample code for http://www.aspose.com/Community/Forums/thread/107477.aspx
	''' 
	''' The scenario is as follows:
	''' 
	''' 1. The customer has a DOC file with footnotes.
	''' 
	''' 2. The customer uses Aspose.Words to convert DOC to HTML. Aspose.Words converts
	''' footnotes and endnotes into hyperlinks. There are two hyperlinks per footnote actually.
	''' One link is "forward" from the main text to the text of the footnote. 
	''' Another is "backward" from the text of the footnote to the main text.
	''' 
	''' 3. The customer uses Aspose.Words to convert HTML back to DOC.
	''' In the current version of Aspose.Words the hyperlinks do not become footnotes,
	''' they just stay as hyperlink fields in the document. The customer wants 
	''' original footnotes to become footnotes during DOC->HTML->DOC roundtrip.
	''' 
	''' This code is a workaround that detects hyperlinks related to footnotes and converts
	''' them into proper footnotes. At some point in the future, this code will not be needed
	''' when Aspose.Words will guarantee footnotes roundtripping.
	''' 
	''' This code demonstrates some useful techniques, such as enumerating over nodes,
	''' getting field code, removing fields etc.
	''' </summary>
	Friend Class Program
		Public Shared Sub Main(ByVal args() As String)
			Dim exeDir As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar
			Dim dataDir As String = New Uri(New Uri(exeDir), "../../Data/").LocalPath

			' Load DOC with footnotes into a document object.
			Dim srcDoc As New Document(Path.Combine(dataDir, "FootnoteSample.doc"))

			' Save to HTML file. Footnotes get converted to hyperlinks.
			Dim saveOptions As New HtmlSaveOptions()
			saveOptions.PrettyFormat = True
			Dim htmlFile As String = Path.Combine(dataDir, "FootnoteSample Out.html")
			srcDoc.Save(htmlFile, saveOptions)

			' Load HTML back into a document object. 
			' In the current version of Aspose.Words hyperlinks do not become footnotes again,
			' they become regular hyperlinks.
			Dim dstDoc As New Document(htmlFile)

			' You can open this document in MS Word and see there are no footnotes, just hyperlinks.
			dstDoc.Save(Path.Combine(dataDir, "FootnoteSample Out1.doc"))

			' This is the workaround method I'm suggesting. It will recognize hyperlinks that
			' should become footnotes and convert them into footnotes.
			ConvertHyperlinksToFootnotes(dstDoc)

			' You can open this document in MS Word and see the footnotes are as expected.
			dstDoc.Save(Path.Combine(dataDir, "FootnoteSample Out2.doc"))
		End Sub

		''' <summary>
		''' A "workaround" method that you can use after DOC->HTML->DOC conversion of a document
		''' with footnotes. Will make sure that original DOC footnotes will still be footnotes in 
		''' the final DOC file.
		''' </summary>
		Friend Shared Sub ConvertHyperlinksToFootnotes(ByVal doc As Document)
			' When processing HYPERLINK fields we will remove them (convert to footnotes).
			' Since it is not a good thing to delete nodes while iterating over a collection, 
			' we will collect the nodes during the first pass and delete them during the second.
			'
			' These collections contain HYPERLINK field starts of footnotes and endnotes in the main document.
			Dim ftnFieldStarts As New Hashtable()
			Dim ednFieldStarts As New Hashtable()
			' These collections contain HYPERLINK field starts of footnotes and endnotes themselves.
			Dim ftnRefFieldStarts As New Hashtable()
			Dim ednRefFieldStarts As New Hashtable()

			' Collect all the nodes into arrays before we start deleting them.
			CollectFieldStarts(doc, ftnFieldStarts, ednFieldStarts, ftnRefFieldStarts, ednRefFieldStarts)

			' Remove the HR shapes that separate footnotes and endnotes from the main text.
			RemoveHorizontalLine(ftnRefFieldStarts)
			RemoveHorizontalLine(ednRefFieldStarts)

			' Convert the HYPERLINK fields into proper footnotes and endnotes.
			ConvertFieldsToNotes(ftnFieldStarts, ftnRefFieldStarts, FootnoteType.Footnote)
			ConvertFieldsToNotes(ednFieldStarts, ednRefFieldStarts, FootnoteType.Endnote)
		End Sub

		''' <summary>
		''' Collects field start nodes of HYPERLINK fields related to footnotes and endnotes.
		''' </summary>
		''' <param name="doc">The document to process.</param>
		''' <param name="ftnFieldStarts">Starts of HYPERLINK fields that represent footnotes will be returned here.</param>
		''' <param name="ednFieldStarts">Start of HYPERLINK fields that represent endnotes will be returned here.</param>
		''' <param name="ftnRefFieldStarts">Starts of HYPERLINK fields that are back-links to footnotes will be returned here.</param>
		''' <param name="ednRefFieldStarts">Starts of HYPERLINK fields that are back-links to endnotes will be returned here.</param>
		Private Shared Sub CollectFieldStarts(ByVal doc As Document, ByVal ftnFieldStarts As Hashtable, ByVal ednFieldStarts As Hashtable, ByVal ftnRefFieldStarts As Hashtable, ByVal ednRefFieldStarts As Hashtable)
			' This regex parses the "command" which we use to determine the footnote/endnote type
			' and the id.
			Dim regex As New Regex("HYPERLINK \\l \""(?<cmd>(_ftn|_edn|_ftnref|_ednref))(?<id>[0-9]+)\""")

			' We need to process all HYPERLINK fields. Therefore select all field starts.
			Dim fieldStarts As NodeCollection = doc.GetChildNodes(NodeType.FieldStart, True)
			For Each fieldStart As FieldStart In fieldStarts
				If fieldStart.FieldType = FieldType.FieldHyperlink Then
					' The field is a hyperlink, lets analyze the field code.
					Dim fieldCode As String = GetFieldCode(fieldStart)

					Dim match As Match = regex.Match(fieldCode)
					Dim cmd As String = match.Groups("cmd").Value
					Dim id As String = match.Groups("id").Value

					Select Case cmd
						Case "_ftn"
							' Field is HYPERLINK \l "_ftn1". It is a footnote in the main document.
							ftnFieldStarts.Add(Integer.Parse(id), fieldStart)
						Case "_edn"
							ednFieldStarts.Add(Integer.Parse(id), fieldStart)
						Case "_ftnref"
							' Field is HYPERLINK \l "_ftnref1". It is a back-link to the footnote in 
							' the main document. The parent paragraph contains the text of the footnote.
							ftnRefFieldStarts.Add(Integer.Parse(id), fieldStart)
						Case "_ednref"
							ednRefFieldStarts.Add(Integer.Parse(id), fieldStart)
						Case Else
							' Do nothing.
					End Select
				End If
			Next fieldStart
		End Sub

		''' <summary>
		''' A simplistic method to get the field code as a string.
		''' Goes trough all Run nodes after the field start and concatenates their text.
		''' </summary>
		Private Shared Function GetFieldCode(ByVal fieldStart As FieldStart) As String
			Dim fieldCode As New StringBuilder()
			Dim curNode As Node = fieldStart.NextSibling
			Do While TypeOf curNode Is Run
				fieldCode.Append(curNode.GetText())
				curNode = curNode.NextSibling
			Loop
			Return fieldCode.ToString()
		End Function

		''' <summary>
		''' Performs the actual conversion of HYPERLINK fields into footnotes/endnote.
		''' </summary>
		''' <param name="noteFieldStarts">The starts of hyperlink fields in the main document.</param>
		''' <param name="refNoteFieldStarts">The starts of back-link hyperlink fields in the footnotes.</param>
		''' <param name="noteType">Specifies whether we are processing footnotes or endnotes.</param>
		Private Shared Sub ConvertFieldsToNotes(ByVal noteFieldStarts As Hashtable, ByVal refNoteFieldStarts As Hashtable, ByVal noteType As FootnoteType)
			For Each entry As DictionaryEntry In noteFieldStarts
				' Footnote/endnote id is stored in the key.
				Dim id As Integer = CInt(Fix(entry.Key))
				Dim noteFieldStart As FieldStart = CType(entry.Value, FieldStart)
				' Using the id we can retrieve the field start of the back-link field.
				Dim refNoteFieldStart As FieldStart = CType(refNoteFieldStarts(id), FieldStart)

				ConvertFieldToNote(noteFieldStart, refNoteFieldStart, noteType)
			Next entry
		End Sub

		''' <summary>
		''' Performs the actual task of converting one HYPERLINK into footnote or endnote.
		''' </summary>
		''' <param name="noteFieldStart">The start of the hyperlink field in the main document.</param>
		''' <param name="refNoteFieldStart">The start of the back-link hyperlink field in the footnote.</param>
		''' <param name="noteType">Specifies whether we are processing footnotes or endnotes.</param>
		Private Shared Sub ConvertFieldToNote(ByVal noteFieldStart As FieldStart, ByVal refNoteFieldStart As FieldStart, ByVal noteType As FootnoteType)
			' This is the paragraph that contains the text of the footnote.
			Dim oldNotePara As Paragraph = refNoteFieldStart.ParentParagraph

			' Delete the hyperlink field from the text of the footnote because we don't need it anymore.
			DeleteField(refNoteFieldStart)

			' Use document builder to move to the place in the main document where the footnote
			' should be and insert a proper footnote.
			Dim builder As New DocumentBuilder(CType(noteFieldStart.Document, Document))
			builder.MoveTo(noteFieldStart)
			Dim note As Footnote = builder.InsertFootnote(noteType, "")

			' Move all content from the old footnote paragraphs into the new.
			Dim newNotePara As Paragraph = note.FirstParagraph
			Dim curNode As Node = oldNotePara.FirstChild
			Do While curNode IsNot Nothing
				Dim nextNode As Node = curNode.NextSibling
				newNotePara.AppendChild(curNode)
				curNode = nextNode
			Loop

			' Delete the old paragraph that represented the footnote. 
			oldNotePara.Remove()

			' Remove the hyperlink field from the main text to the footnote.
			DeleteField(noteFieldStart)
		End Sub

		''' <summary>
		''' A simplistic method to delete all nodes of a field given a field start node.
		''' </summary>
		Private Shared Sub DeleteField(ByVal fieldStart As FieldStart)
			Dim curNode As Node = fieldStart
			Do While curNode.NodeType <> NodeType.FieldEnd
				Dim nextNode As Node = curNode.NextSibling
				curNode.Remove()
				curNode = nextNode
			Loop
			curNode.Remove()
		End Sub

		''' <summary>
		''' There is an HR (horizontal rule) shape in a separate paragraph just before
		''' the first footnote and first endnote in a document imported from HTML.
		''' This method deletes the paragraph and the HR shape.
		''' </summary>
		Private Shared Sub RemoveHorizontalLine(ByVal noteRefFieldStarts As Hashtable)
			' Footnote and endnote ids start from 1. Therefore we can get the first note.
			Dim noteFieldStart As FieldStart = CType(noteRefFieldStarts(1), FieldStart)
			' This is the paragraph that contains the first footnote.
			Dim notePara As Paragraph = noteFieldStart.ParentParagraph
			' This is the previous paragraph that contains the HR shape. Delete the paragraph.
			Dim hrPara As Paragraph = CType(notePara.PreviousSibling, Paragraph)
			hrPara.Remove()
		End Sub
	End Class
End Namespace
