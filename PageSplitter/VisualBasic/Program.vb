'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.IO

Imports Aspose.Words
Imports Aspose.Words.Layout

Namespace PageSplitter
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			Dim dataDir As String = Path.GetFullPath("../../Data/")
			SplitAllDocumentsToPages(dataDir)
		End Sub

		Public Shared Sub SplitDocumentToPages(ByVal docName As String)
			Dim folderName As String = Path.GetDirectoryName(docName)
			Dim fileName As String = Path.GetFileNameWithoutExtension(docName)
			Dim extensionName As String = Path.GetExtension(docName)
			Dim outFolder As String = Path.Combine(folderName, "Out")

			Console.WriteLine("Processing document: " & fileName & extensionName)

			Dim doc As New Document(docName)

			' Create and attach collector to the document before page layout is built.
			Dim layoutCollector As New LayoutCollector(doc)

			' This will build layout model and collect necessary information.
			doc.UpdatePageLayout()

			' Split nodes in the document into separate pages.
			Dim splitter As New DocumentPageSplitter(layoutCollector)

			' Save each page to the disk as a separate document.
			For page As Integer = 1 To doc.PageCount
				Dim pageDoc As Document = splitter.GetDocumentOfPage(page)
				pageDoc.Save(Path.Combine(outFolder, String.Format("{0} - page{1} Out{2}", fileName, page, extensionName)))
			Next page

			' Detach the collector from the document.
			layoutCollector.Document = Nothing
		End Sub

		Public Shared Sub SplitAllDocumentsToPages(ByVal folderName As String)
			Dim fileNames() As String = Directory.GetFiles(folderName, "*.doc?", SearchOption.TopDirectoryOnly)

			For Each fileName As String In fileNames
				SplitDocumentToPages(fileName)
			Next fileName
		End Sub
	End Class
End Namespace
