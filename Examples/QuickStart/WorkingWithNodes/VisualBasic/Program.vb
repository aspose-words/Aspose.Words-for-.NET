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

Imports Aspose.Words

Namespace WorkingWithNodesExample
	Public Class Program
		Public Shared Sub Main()
			' Create a new document.
			Dim doc As New Document()

			' Creates and adds a paragraph node to the document.
			Dim para As New Paragraph(doc)

			' Typed access to the last section of the document.
			Dim section As Section = doc.LastSection
			section.Body.AppendChild(para)

			' Next print the node type of one of the nodes in the document.
			Dim nodeType As NodeType = doc.FirstSection.Body.NodeType

			Console.WriteLine("NodeType: " & Node.NodeTypeToString(nodeType))
		End Sub
	End Class
End Namespace