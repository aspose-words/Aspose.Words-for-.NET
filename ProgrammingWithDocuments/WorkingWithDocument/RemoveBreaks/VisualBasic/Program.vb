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
Imports System.Reflection

Imports Aspose.Words

Namespace RemoveBreaksExample
	Public Class Program
		Public Shared Sub Main()
			' The sample infrastructure.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Open the document.
			Dim doc As New Document(dataDir & "TestFile.doc")

			' Remove the page and section breaks from the document.
			' In Aspose.Words section breaks are represented as separate Section nodes in the document.
			' To remove these separate sections the sections are combined.
			RemovePageBreaks(doc)
			RemoveSectionBreaks(doc)

			' Save the document.
			doc.Save(dataDir & "TestFile Out.doc")
		End Sub

		'ExStart
		'ExFor:ControlChar.PageBreak
		'ExId:RemoveBreaks_Pages
		'ExSummary:Removes all page breaks from the document.
		Private Shared Sub RemovePageBreaks(ByVal doc As Document)
			' Retrieve all paragraphs in the document.
			Dim paragraphs As NodeCollection = doc.GetChildNodes(NodeType.Paragraph, True)

			' Iterate through all paragraphs
			For Each para As Paragraph In paragraphs
				' If the paragraph has a page break before set then clear it.
				If para.ParagraphFormat.PageBreakBefore Then
					para.ParagraphFormat.PageBreakBefore = False
				End If

				' Check all runs in the paragraph for page breaks and remove them.
				For Each run As Run In para.Runs
					If run.Text.Contains(ControlChar.PageBreak) Then
						run.Text = run.Text.Replace(ControlChar.PageBreak, String.Empty)
					End If
				Next run

			Next para

		End Sub
		'ExEnd


		'ExStart
		'ExId:RemoveBreaks_Sections
		'ExSummary:Combines all sections in the document into one.
		Private Shared Sub RemoveSectionBreaks(ByVal doc As Document)
			' Loop through all sections starting from the section that precedes the last one 
			' and moving to the first section.
			For i As Integer = doc.Sections.Count - 2 To 0 Step -1
				' Copy the content of the current section to the beginning of the last section.
				doc.LastSection.PrependContent(doc.Sections(i))
				' Remove the copied section.
				doc.Sections(i).Remove()
			Next i
		End Sub
		'ExEnd
	End Class
End Namespace