'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports Aspose.Words
Imports NUnit.Framework

Namespace Examples
	<TestFixture> _
	Public Class ExStyles
		Inherits ExBase
		<Test> _
		Public Sub GetStyles()
			'ExStart
			'ExFor:DocumentBase.Styles
			'ExFor:Style.Name
			'ExId:GetStyles
			'ExSummary:Shows how to get access to the collection of styles defined in the document.
			Dim doc As New Document()
			Dim styles As StyleCollection = doc.Styles

			For Each style As Style In styles
				Console.WriteLine(style.Name)
			Next style
			'ExEnd
		End Sub

		<Test> _
		Public Sub SetAllStyles()
			'ExStart
			'ExFor:Style.Font
			'ExFor:Style
			'ExSummary:Shows how to change the font formatting of all styles in a document.
			Dim doc As New Document()
			For Each style As Style In doc.Styles
				If style.Font IsNot Nothing Then
					style.Font.ClearFormatting()
					style.Font.Size = 20
					style.Font.Name = "Arial"
				End If
			Next style
			'ExEnd
		End Sub

		<Test> _
		Public Sub ChangeStyleOfTOCLevel()
			Dim doc As New Document()
			'ExStart
			'ExId:ChangeTOCStyle
			'ExSummary:Changes a formatting property used in the first level TOC style.
			' Retrieve the style used for the first level of the TOC and change the formatting of the style.
			doc.Styles(StyleIdentifier.Toc1).Font.Bold = True
			'ExEnd
		End Sub

		<Test> _
		Public Sub ChangeTOCTabStops()
			'ExStart
			'ExFor:TabStop
			'ExFor:ParagraphFormat.TabStops
			'ExFor:Style.StyleIdentifier
			'ExFor:TabStopCollection.RemoveByPosition
			'ExFor:TabStop.Alignment
			'ExFor:TabStop.Position
			'ExFor:TabStop.Leader
			'ExId:ChangeTOCTabStops
			'ExSummary:Shows how to modify the position of the right tab stop in TOC related paragraphs.
			Dim doc As New Document(MyDir & "Document.TableOfContents.doc")

			' Iterate through all paragraphs in the document
			For Each para As Paragraph In doc.GetChildNodes(NodeType.Paragraph, True)
				' Check if this paragraph is formatted using the TOC result based styles. This is any style between TOC and TOC9.
				If para.ParagraphFormat.Style.StyleIdentifier >= StyleIdentifier.Toc1 AndAlso para.ParagraphFormat.Style.StyleIdentifier <= StyleIdentifier.Toc9 Then
					' Get the first tab used in this paragraph, this should be the tab used to align the page numbers.
					Dim tab As TabStop = para.ParagraphFormat.TabStops(0)
					' Remove the old tab from the collection.
					para.ParagraphFormat.TabStops.RemoveByPosition(tab.Position)
					' Insert a new tab using the same properties but at a modified position. 
					' We could also change the separators used (dots) by passing a different Leader type
					para.ParagraphFormat.TabStops.Add(tab.Position - 50, tab.Alignment, tab.Leader)
				End If
			Next para

			doc.Save(MyDir & "Document.TableOfContentsTabStops Out.doc")
			'ExEnd
		End Sub
	End Class
End Namespace
