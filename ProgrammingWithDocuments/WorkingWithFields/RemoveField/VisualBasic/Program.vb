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
Imports Aspose.Words.Fields

Namespace RemoveFieldExample
	Public Class Program
		Public Shared Sub Main()
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			Dim doc As New Document(dataDir & "RemoveField.doc")

			'ExStart
			'ExFor:Field.Remove
			'ExId:DocumentBuilder_RemoveField
			'ExSummary:Removes a field from the document.
			Dim field As Field = doc.Range.Fields(0)
			' Calling this method completely removes the field from the document.
			field.Remove()
			'ExEnd

		End Sub
	End Class
End Namespace