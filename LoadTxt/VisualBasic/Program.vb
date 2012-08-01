'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////
'ExStart
'ExId:LoadTxt
'ExSummary:Loads a plain text file into an Aspose.Words.Document object.

Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Reflection
Imports System.Text

Imports Aspose.Words

Namespace LoadTxt
	Friend Class Program
		Public Shared Sub Main(ByVal args() As String)
			' Sample infrastructure.
			Dim exeDir As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar
			Dim dataDir As String = New Uri(New Uri(exeDir), "../../Data/").LocalPath

			' The encoding of the text file is automatically detected.
			Dim doc As New Document(dataDir & "LoadTxt.txt")

			' Save as any Aspose.Words supported format, such as DOCX.
			doc.Save(dataDir & "LoadTxt Out.docx")
		End Sub
	End Class
End Namespace
'ExEnd
