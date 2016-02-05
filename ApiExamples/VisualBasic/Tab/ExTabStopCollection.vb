' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports NUnit.Framework


Namespace ApiExamples.Tab
	<TestFixture> _
	Public Class ExTabStopCollection
		Inherits ApiExampleBase
		<Test> _
		Public Sub ClearEx()
			'ExStart
			'ExFor:TabStopCollection.Clear
			'ExSummary:Shows how to remove all tab stops from a document.
			Dim doc As New Aspose.Words.Document(MyDir & "Document.TableOfContents.doc")

			' Clear all tab stops from every paragraph.
			For Each para As Aspose.Words.Paragraph In doc.GetChildNodes(Aspose.Words.NodeType.Paragraph, True)
				para.ParagraphFormat.TabStops.Clear()
			Next para

			doc.Save(MyDir & "Document.AllTabStopsRemoved Out.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub AddEx()
			'ExStart
			'ExFor:TabStopCollection.Add(TabStop)
			'ExFor:TabStopCollection.Add(Double, TabAlignment, TabLeader)
			'ExSummary:Shows how to create tab stops and add them to a document.
			Dim doc As New Aspose.Words.Document(MyDir & "Document.doc")
			Dim paragraph As Aspose.Words.Paragraph = CType(doc.GetChild(Aspose.Words.NodeType.Paragraph, 0, True), Aspose.Words.Paragraph)

			' Create a TabStop object and add it to the document.
			Dim tabStop As New Aspose.Words.TabStop(Aspose.Words.ConvertUtil.InchToPoint(3), Aspose.Words.TabAlignment.Left, Aspose.Words.TabLeader.Dashes)
			paragraph.ParagraphFormat.TabStops.Add(tabStop)

			' Add a tab stop without explicitly creating new TabStop objects.
			paragraph.ParagraphFormat.TabStops.Add(Aspose.Words.ConvertUtil.MillimeterToPoint(100), Aspose.Words.TabAlignment.Left, Aspose.Words.TabLeader.Dashes)

			' Add tab stops at 5 cm to all paragraphs.
			For Each para As Aspose.Words.Paragraph In doc.GetChildNodes(Aspose.Words.NodeType.Paragraph, True)
				para.ParagraphFormat.TabStops.Add(Aspose.Words.ConvertUtil.MillimeterToPoint(50), Aspose.Words.TabAlignment.Left, Aspose.Words.TabLeader.Dashes)
			Next para

			doc.Save(MyDir & "Document.AddedTabStops Out.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub RemoveByIndexEx()
			'ExStart
			'ExFor:TabStopCollection.RemoveByIndex
			'ExSummary:Shows how to select a tab stop in a document by its index and remove it.
			Dim doc As New Aspose.Words.Document(MyDir & "Document.doc")
			Dim paragraph As Aspose.Words.Paragraph = CType(doc.GetChild(Aspose.Words.NodeType.Paragraph, 0, True), Aspose.Words.Paragraph)

			paragraph.ParagraphFormat.TabStops.Add(Aspose.Words.ConvertUtil.MillimeterToPoint(30), Aspose.Words.TabAlignment.Left, Aspose.Words.TabLeader.Dashes)
			paragraph.ParagraphFormat.TabStops.Add(Aspose.Words.ConvertUtil.MillimeterToPoint(60), Aspose.Words.TabAlignment.Left, Aspose.Words.TabLeader.Dashes)

			' Tab stop placed at 30 mm is removed
			paragraph.ParagraphFormat.TabStops.RemoveByIndex(0)

			Console.WriteLine(paragraph.ParagraphFormat.TabStops.Count)

			doc.Save(MyDir & "Document.RemovedTabStopsByIndex Out.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub GetPositionByIndexEx()
			'ExStart
			'ExFor:TabStopCollection.GetPositionByIndex
			'ExSummary:Shows how to find a tab stop by it's index and get its position.
			Dim doc As New Aspose.Words.Document(MyDir & "Document.doc")
			Dim paragraph As Aspose.Words.Paragraph = CType(doc.GetChild(Aspose.Words.NodeType.Paragraph, 0, True), Aspose.Words.Paragraph)

			paragraph.ParagraphFormat.TabStops.Add(Aspose.Words.ConvertUtil.MillimeterToPoint(30), Aspose.Words.TabAlignment.Left, Aspose.Words.TabLeader.Dashes)
			paragraph.ParagraphFormat.TabStops.Add(Aspose.Words.ConvertUtil.MillimeterToPoint(60), Aspose.Words.TabAlignment.Left, Aspose.Words.TabLeader.Dashes)

			Console.WriteLine("Tab stop at index {0} of the first paragraph is at {1} points.", 1, paragraph.ParagraphFormat.TabStops.GetPositionByIndex(1))
			'ExEnd
		End Sub

		<Test> _
		Public Sub GetIndexByPositionEx()
			'ExStart
			'ExFor:TabStopCollection.GetIndexByPosition
			'ExSummary:Shows how to look up a position to see if a tab stop exists there, and if so, obtain its index.
			Dim doc As New Aspose.Words.Document(MyDir & "Document.doc")
			Dim paragraph As Aspose.Words.Paragraph = CType(doc.GetChild(Aspose.Words.NodeType.Paragraph, 0, True), Aspose.Words.Paragraph)

			paragraph.ParagraphFormat.TabStops.Add(Aspose.Words.ConvertUtil.MillimeterToPoint(30), Aspose.Words.TabAlignment.Left, Aspose.Words.TabLeader.Dashes)

			' An output of -1 signifies that there is no tab stop at that position.
			Console.WriteLine(paragraph.ParagraphFormat.TabStops.GetIndexByPosition(Aspose.Words.ConvertUtil.MillimeterToPoint(30))) ' 0
			Console.WriteLine(paragraph.ParagraphFormat.TabStops.GetIndexByPosition(Aspose.Words.ConvertUtil.MillimeterToPoint(60))) ' -1
			'ExEnd
		End Sub
	End Class
End Namespace
