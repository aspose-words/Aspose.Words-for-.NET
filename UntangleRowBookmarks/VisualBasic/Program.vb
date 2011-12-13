'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
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
Imports Aspose.Words.Tables

Namespace UntangleRowBookmarks
	Public Class Program
		''' <summary>
		''' The main entry point for the application.
		''' </summary>
		Public Shared Sub Main(ByVal args() As String)
			Dim exeDir As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar
			Dim dataDir As String = New Uri(New Uri(exeDir), "../../Data/").LocalPath

			' Load a document.
			Dim doc As New Document(dataDir & "TestDefect1352.doc")

			' This perform the custom task of putting the row bookmark ends into the same row with the bookmark starts.
			UntangleRowBookmarks(doc)

			' Now we can easily delete rows by a bookmark without damaging any other row's bookmarks.
			DeleteRowByBookmark(doc, "ROW2")

			' This is just to check that the other bookmark was not damaged.
			If doc.Range.Bookmarks("ROW1").BookmarkEnd Is Nothing Then
				Throw New Exception("Wrong, the end of the bookmark was deleted.")
			End If

			' Save the finished document.
			doc.Save(dataDir & "TestDefect1352 Out.doc")
		End Sub

		Private Shared Sub UntangleRowBookmarks(ByVal doc As Document)
			For Each bookmark As Bookmark In doc.Range.Bookmarks
				' Get the parent row of both the bookmark and bookmark end node.
				Dim row1 As Row = CType(bookmark.BookmarkStart.GetAncestor(GetType(Row)), Row)
				Dim row2 As Row = CType(bookmark.BookmarkEnd.GetAncestor(GetType(Row)), Row)

				' If both rows are found okay and the bookmark start and end are contained
				' in adjacent rows, then just move the bookmark end node to the end
				' of the last paragraph in the last cell of the top row.
				If (row1 IsNot Nothing) AndAlso (row2 IsNot Nothing) AndAlso (row1.NextSibling Is row2) Then
					row1.LastCell.LastParagraph.AppendChild(bookmark.BookmarkEnd)
				End If
			Next bookmark
		End Sub

		Private Shared Sub DeleteRowByBookmark(ByVal doc As Document, ByVal bookmarkName As String)
			' Find the bookmark in the document. Exit if cannot find it.
			Dim bookmark As Bookmark = doc.Range.Bookmarks(bookmarkName)
			If bookmark Is Nothing Then
				Return
			End If

			' Get the parent row of the bookmark. Exit if the bookmark is not in a row.
			Dim row As Row = CType(bookmark.BookmarkStart.GetAncestor(GetType(Row)), Row)
			If row Is Nothing Then
				Return
			End If

			' Remove the row.
			row.Remove()
		End Sub
	End Class
End Namespace
