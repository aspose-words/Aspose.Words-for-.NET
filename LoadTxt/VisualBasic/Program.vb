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

			' This object will help us generate the document.
			Dim builder As New DocumentBuilder()

			' You might need to specify a different encoding depending on your plain text files.
			Using reader As New StreamReader(dataDir & "LoadTxt.txt", Encoding.UTF8)
				' Read plain text "lines" and convert them into paragraphs in the document.
				Dim line As String = Nothing
				line = reader.ReadLine()
				Do While line IsNot Nothing
					builder.Writeln(line)
					line = reader.ReadLine()
				Loop
			End Using

			' Save in any Aspose.Words supported format.
			builder.Document.Save(dataDir & "LoadTxt Out.docx")
			builder.Document.Save(dataDir & "LoadTxt Out.html")
		End Sub
	End Class
End Namespace
'ExEnd
