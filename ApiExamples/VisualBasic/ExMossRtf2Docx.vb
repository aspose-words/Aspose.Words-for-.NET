' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports Aspose.Words

Namespace ApiExamples
	Public Class ExMossRtf2Docx
		'ExStart
		'ExId:MossRtf2Docx
		'ExSummary:Converts an RTF document to OOXML.
		Public Shared Sub ConvertRtfToDocx(ByVal inFileName As String, ByVal outFileName As String)
			' Load an RTF file into Aspose.Words.
			Dim doc As New Document(inFileName)

			' Save the document in the OOXML format.
			doc.Save(outFileName, SaveFormat.Docx)
		End Sub
		'ExEnd
	End Class
End Namespace
