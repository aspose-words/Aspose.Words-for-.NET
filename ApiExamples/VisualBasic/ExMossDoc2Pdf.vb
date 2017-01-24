' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

'ExStart
'ExId:MossDoc2Pdf
'ExSummary:The following is the complete code of the document converter.


Imports Microsoft.VisualBasic
Imports System
Imports System.IO

Imports Aspose.Words
Imports Aspose.Words.Saving

Namespace ApiExamples
	''' <summary>
	''' DOC2PDF document converter for SharePoint.
	''' Uses Aspose.Words to perform the conversion.
	''' </summary>
	Public Class ExMossDoc2Pdf
		''' <summary>
		''' The main entry point for the application.
		''' </summary>
		<STAThread> _
		Public Shared Sub Main(ByVal args() As String)
			' Although SharePoint passes "-log <filename>" to us and we are
			' supposed to log there, for the sake of simplicity, we will use 
			' our own hard coded path to the log file.
			' 
			' Make sure there are permissions to write into this folder.
			' The document converter will be called under the document 
			' conversion account (not sure what name), so for testing purposes 
			' I would give the Users group write permissions into this folder.
			gLog = New StreamWriter("C:\Aspose2Pdf\log.txt", True)

			Try
				gLog.WriteLine(DateTime.Now.ToString() & " Started")
				gLog.WriteLine(Environment.CommandLine)

				ParseCommandLine(args)

				' Uncomment the code below when you have purchased a licenses for Aspose.Words.
				'
				' You need to deploy the license in the same folder as your 
				' executable, alternatively you can add the license file as an 
				' embedded resource to your project.
				'
				' // Set license for Aspose.Words.
				' Aspose.Words.License wordsLicense = new Aspose.Words.License();
				' wordsLicense.SetLicense("Aspose.Total.lic");

				ConvertDoc2Pdf(gInFileName, gOutFileName)
			Catch e As Exception
				gLog.WriteLine(e.Message)
				Environment.ExitCode = 100
			Finally
				gLog.Close()
			End Try
		End Sub

		Private Shared Sub ParseCommandLine(ByVal args() As String)
			Dim i As Integer = 0
			Do While i < args.Length
				Dim s As String = args(i)
				Select Case s.ToLower()
					Case "-in"
						i += 1
						gInFileName = args(i)
					Case "-out"
						i += 1
						gOutFileName = args(i)
					Case "-config"
						' Skip the name of the config file and do nothing.
						i += 1
					Case "-log"
						' Skip the name of the log file and do nothing.
						i += 1
					Case Else
						Throw New Exception("Unknown command line argument: " & s)
				End Select
				i += 1
			Loop
		End Sub

		Private Shared Sub ConvertDoc2Pdf(ByVal inFileName As String, ByVal outFileName As String)
			' You can load not only DOC here, but any format supported by
			' Aspose.Words: DOC, DOCX, RTF, WordML, HTML, MHTML, ODT etc.
			Dim doc As New Document(inFileName)

			doc.Save(outFileName, New PdfSaveOptions())
		End Sub

		Private Shared gInFileName As String
		Private Shared gOutFileName As String
		Private Shared gLog As StreamWriter
	End Class
End Namespace
'ExEnd


