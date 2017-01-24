' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////



Imports Microsoft.VisualBasic
Imports NUnit.Framework
Imports System.IO

Imports Aspose.Words
Imports Aspose.Pdf.Facades
Imports Aspose.Words.Saving

Namespace ApiExamples
	<TestFixture> _
	Public Class ExBookmarks
		Inherits ApiExampleBase
		<Test> _
		Public Sub BookmarkNameAndText()
			'ExStart
			'ExFor:Bookmark
			'ExFor:Bookmark.Name
			'ExFor:Bookmark.Text
			'ExFor:Range.Bookmarks
			'ExId:BookmarksGetNameSetText
			'ExSummary:Shows how to get or set bookmark name and text.
			Dim doc As New Document(MyDir & "Bookmark.doc")

			' Use the indexer of the Bookmarks collection to obtain the desired bookmark.
			Dim bookmark As Aspose.Words.Bookmark = doc.Range.Bookmarks("MyBookmark")

			' Get the name and text of the bookmark.
			Dim name As String = bookmark.Name
			Dim text As String = bookmark.Text

			' Set the name and text of the bookmark.
			bookmark.Name = "RenamedBookmark"
			bookmark.Text = "This is a new bookmarked text."
			'ExEnd

			Assert.AreEqual("MyBookmark", name)
			Assert.AreEqual("This is a bookmarked text.", text)
		End Sub

		<Test> _
		Public Sub BookmarkRemove()
			'ExStart
			'ExFor:Bookmark.Remove
			'ExSummary:Shows how to remove a particular bookmark from a document.
			Dim doc As New Document(MyDir & "Bookmark.doc")

			' Use the indexer of the Bookmarks collection to obtain the desired bookmark.
			Dim bookmark As Aspose.Words.Bookmark = doc.Range.Bookmarks("MyBookmark")

			' Remove the bookmark. The bookmarked text is not deleted.
			bookmark.Remove()
			'ExEnd

			' Verify that the bookmarks were removed from the document.
			Assert.AreEqual(0, doc.Range.Bookmarks.Count)
		End Sub

		<Test> _
		Public Sub ClearBookmarks()
			'ExStart
			'ExFor:BookmarkCollection.Clear
			'ExSummary:Shows how to remove all bookmarks from a document.
			Dim doc As New Document(MyDir & "Bookmark.doc")
			doc.Range.Bookmarks.Clear()
			'ExEnd

			' Verify that the bookmarks were removed
			Assert.AreEqual(0, doc.Range.Bookmarks.Count)
		End Sub

		<Test> _
		Public Sub AccessBookmarks()
			'ExStart
			'ExFor:BookmarkCollection
			'ExFor:BookmarkCollection.Item(Int32)
			'ExFor:BookmarkCollection.Item(String)
			'ExId:BookmarksAccess
			'ExSummary:Shows how to obtain bookmarks from a bookmark collection.
			Dim doc As New Document(MyDir & "Bookmarks.doc")

			' By index.
			Dim bookmark1 As Aspose.Words.Bookmark = doc.Range.Bookmarks(0)

			' By name.
			Dim bookmark2 As Aspose.Words.Bookmark = doc.Range.Bookmarks("Bookmark2")
			'ExEnd
		End Sub

		<Test> _
		Public Sub BookmarkCollectionRemove()
			'ExStart
			'ExFor:BookmarkCollection.Remove(Bookmark)
			'ExFor:BookmarkCollection.Remove(String)
			'ExFor:BookmarkCollection.RemoveAt
			'ExSummary:Demonstrates different methods of removing bookmarks from a document.
			Dim doc As New Document(MyDir & "Bookmarks.doc")
			' Remove a particular bookmark from the document.
			Dim bookmark As Aspose.Words.Bookmark = doc.Range.Bookmarks(0)
			doc.Range.Bookmarks.Remove(bookmark)

			' Remove a bookmark by specified name.
			doc.Range.Bookmarks.Remove("Bookmark2")

			' Remove a bookmark at the specified index.
			doc.Range.Bookmarks.RemoveAt(0)
			'ExEnd

			Assert.AreEqual(0, doc.Range.Bookmarks.Count)
		End Sub

		<Test> _
		Public Sub BookmarksInsertBookmarkWithDocumentBuilder()
			'ExStart
			'ExId:BookmarksInsertBookmark
			'ExSummary:Shows how to create a new bookmark.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			builder.StartBookmark("MyBookmark")
			builder.Writeln("Text inside a bookmark.")
			builder.EndBookmark("MyBookmark")
			'ExEnd
		End Sub

		<Test> _
		Public Sub GetBookmarkCount()
			'ExStart
			'ExFor:BookmarkCollection.Count
			'ExSummary:Shows how to count the number of bookmarks in a document.
			Dim doc As New Document(MyDir & "Bookmark.doc")

			Dim count As Integer = doc.Range.Bookmarks.Count
			'ExEnd

			Assert.AreEqual(1, count)
		End Sub

		<Test> _
		Public Sub CreateBookmarkWithNodes()
			'ExStart
			'ExFor:BookmarkStart
			'ExFor:BookmarkStart.#ctor
			'ExFor:BookmarkEnd
			'ExFor:BookmarkEnd.#ctor
			'ExSummary:Shows how to create a bookmark by inserting bookmark start and end nodes.
			Dim doc As New Document()

			' An empty document has just one empty paragraph by default.
			Dim p As Paragraph = doc.FirstSection.Body.FirstParagraph

			p.AppendChild(New Run(doc, "Text before bookmark. "))

			p.AppendChild(New BookmarkStart(doc, "My bookmark"))
			p.AppendChild(New Run(doc, "Text inside bookmark. "))
			p.AppendChild(New BookmarkEnd(doc, "My bookmark"))

			p.AppendChild(New Run(doc, "Text after bookmark."))

			doc.Save(MyDir & "\Artifacts\Bookmarks.CreateBookmarkWithNodes.doc")

			Assert.AreEqual(doc.Range.Bookmarks("My bookmark").Text, "Text inside bookmark. ")
			'ExEnd
		End Sub

		<Test, TestCase(SaveFormat.Pdf), TestCase(SaveFormat.Xps), TestCase(SaveFormat.Swf)> _
		Public Sub AddBookmarkWithWhiteSpaces(ByVal saveFormat As SaveFormat)
			Dim doc As New Document()

			InsertBookmarks(doc)

            If saveFormat = saveFormat.Pdf Then
                'Save document with pdf save options
                doc.Save(MyDir & "\Artifacts\Bookmark_WhiteSpaces.pdf", AddBookmarkSaveOptions(saveFormat.Pdf))

                'Bind pdf with Aspose PDF
                Dim bookmarkEditor As New PdfBookmarkEditor()
                bookmarkEditor.BindPdf(MyDir & "Bookmark_WhiteSpaces.pdf")

                'Get all bookmarks from the document
                Dim bookmarks As Bookmarks = bookmarkEditor.ExtractBookmarks()

                Assert.AreEqual(3, bookmarks.Count)

                'Assert that all the bookmarks title are with witespaces
                Assert.AreEqual("My Bookmark", bookmarks(0).Title)
                Assert.AreEqual("Nested Bookmark", bookmarks(1).Title)

                'Assert that the bookmark title without witespaces
                Assert.AreEqual("Bookmark_WithoutWhiteSpaces", bookmarks(2).Title)
            Else
                Dim dstStream As New MemoryStream()
                doc.Save(dstStream, AddBookmarkSaveOptions(saveFormat))

                'Get bookmarks from the document
                Dim bookmarks As BookmarkCollection = doc.Range.Bookmarks

                Assert.AreEqual(3, bookmarks.Count)

                'Assert that all the bookmarks title are with witespaces
                Assert.AreEqual("My Bookmark", bookmarks(0).Name)
                Assert.AreEqual("Nested Bookmark", bookmarks(1).Name)

                'Assert that the bookmark title without witespaces
                Assert.AreEqual("Bookmark_WithoutWhiteSpaces", bookmarks(2).Name)
            End If
		End Sub

		Private Shared Sub InsertBookmarks(ByVal doc As Document)
			Dim builder As New DocumentBuilder(doc)

			builder.StartBookmark("My Bookmark")
			builder.Writeln("Text inside a bookmark.")

			builder.StartBookmark("Nested Bookmark")
			builder.Writeln("Text inside a NestedBookmark.")
			builder.EndBookmark("Nested Bookmark")

			builder.Writeln("Text after Nested Bookmark.")
			builder.EndBookmark("My Bookmark")

			builder.StartBookmark("Bookmark_WithoutWhiteSpaces")
			builder.Writeln("Text inside a NestedBookmark.")
			builder.EndBookmark("Bookmark_WithoutWhiteSpaces")
		End Sub

		Private Shared Function AddBookmarkSaveOptions(ByVal saveFormat As SaveFormat) As SaveOptions
			Dim pdfSaveOptions As New PdfSaveOptions()
			Dim xpsSaveOptions As New XpsSaveOptions()
			Dim swfSaveOptions As New SwfSaveOptions()

			Select Case saveFormat
				Case SaveFormat.Pdf

					'Add bookmarks to the document
					pdfSaveOptions.OutlineOptions.BookmarksOutlineLevels.Add("My Bookmark", 1)
					pdfSaveOptions.OutlineOptions.BookmarksOutlineLevels.Add("Nested Bookmark", 2)
					pdfSaveOptions.OutlineOptions.BookmarksOutlineLevels.Add("Bookmark_WithoutWhiteSpaces", 3)

					Return pdfSaveOptions

				Case SaveFormat.Xps

					'Add bookmarks to the document
					xpsSaveOptions.OutlineOptions.BookmarksOutlineLevels.Add("My Bookmark", 1)
					xpsSaveOptions.OutlineOptions.BookmarksOutlineLevels.Add("Nested Bookmark", 2)
					xpsSaveOptions.OutlineOptions.BookmarksOutlineLevels.Add("Bookmark_WithoutWhiteSpaces", 3)

					Return xpsSaveOptions

				Case SaveFormat.Swf

					'Add bookmarks to the document
					swfSaveOptions.OutlineOptions.BookmarksOutlineLevels.Add("My Bookmark", 1)
					swfSaveOptions.OutlineOptions.BookmarksOutlineLevels.Add("Nested Bookmark", 2)
					swfSaveOptions.OutlineOptions.BookmarksOutlineLevels.Add("Bookmark_WithoutWhiteSpaces", 3)

					Return swfSaveOptions
			End Select

			Return Nothing
		End Function
	End Class
End Namespace
