'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Words

Namespace AppendDocumentsExample
	Public Class Program
		Public Shared Sub Main()
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Load the destination and source documents from disk.
			Dim dstDoc As New Document(dataDir & "TestFile.Destination.doc")
			Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

			' Append the source document to the destination document while keeping the original formatting of the source document.
			dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting)

			dstDoc.Save(dataDir & "TestFile Out.docx")
		End Sub
	End Class
End Namespace