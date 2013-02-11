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

Namespace PageNumbersOfNodes
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			Dim dataDir As String = Path.GetFullPath("../../Data/")

			Dim doc As New Document(dataDir & "TestFile.docx")

			' Create and attach collector before the document before page layout is built.
			Dim layoutCollector As New LayoutCollector(doc)

			' This will build layout model and collect necessary information.
			doc.UpdatePageLayout()

			' Print the details of each document node including the page numbers. 
			For Each node As Node In doc.FirstSection.Body.GetChildNodes(NodeType.Any, True)
				Console.WriteLine(" --------- ")
				Console.WriteLine("NodeType:   " & Node.NodeTypeToString(node.NodeType))
				Console.WriteLine("Text:       """ & node.ToString(SaveFormat.Text).Trim() & """")
				Console.WriteLine("Page Start: " & layoutCollector.GetStartPageIndex(node))
				Console.WriteLine("Page End:   " & layoutCollector.GetEndPageIndex(node))
				Console.WriteLine(" --------- ")
				Console.WriteLine()
			Next node

			' Detatch the collector from the document.
			layoutCollector.Document = Nothing

			Console.ReadLine()
		End Sub
	End Class
End Namespace
