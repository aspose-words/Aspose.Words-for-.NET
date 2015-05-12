'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.IO
Imports System.Reflection

Imports Aspose.Words

Namespace CheckFormatExample
	Public Class Program
		Public Shared Sub Main()
			' The sample infrastructure.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")
			Dim supportedDir As String = dataDir & "OutSupported"
			Dim unknownDir As String = dataDir & "OutUnknown"
			Dim encryptedDir As String = dataDir & "OutEncrypted"
			Dim pre97Dir As String = dataDir & "OutPre97"

			'ExStart
			'ExId:CheckFormat_Folder
			'ExSummary:Get the list of all files in the dataDir folder.
			Dim fileList() As String = Directory.GetFiles(dataDir)
			'ExEnd

			'ExStart
			'ExFor:FileFormatInfo
			'ExFor:FileFormatUtil
			'ExFor:FileFormatUtil.DetectFileFormat(String)
			'ExFor:LoadFormat
			'ExId:CheckFormat_Main
			'ExSummary:Check each file in the folder and move it to the appropriate subfolder.
			' Loop through all found files.
			For Each fileName As String In fileList
				' Extract and display the file name without the path.
				Dim nameOnly As String = Path.GetFileName(fileName)
				Console.Write(nameOnly)

				' Check the file format and move the file to the appropriate folder.
				Dim info As FileFormatInfo = FileFormatUtil.DetectFileFormat(fileName)

				' Display the document type.
				Select Case info.LoadFormat
					Case LoadFormat.Doc
						Console.WriteLine(Constants.vbTab & "Microsoft Word 97-2003 document.")
					Case LoadFormat.Dot
						Console.WriteLine(Constants.vbTab & "Microsoft Word 97-2003 template.")
					Case LoadFormat.Docx
						Console.WriteLine(Constants.vbTab & "Office Open XML WordprocessingML Macro-Free Document.")
					Case LoadFormat.Docm
						Console.WriteLine(Constants.vbTab & "Office Open XML WordprocessingML Macro-Enabled Document.")
					Case LoadFormat.Dotx
						Console.WriteLine(Constants.vbTab & "Office Open XML WordprocessingML Macro-Free Template.")
					Case LoadFormat.Dotm
						Console.WriteLine(Constants.vbTab & "Office Open XML WordprocessingML Macro-Enabled Template.")
					Case LoadFormat.FlatOpc
						Console.WriteLine(Constants.vbTab & "Flat OPC document.")
					Case LoadFormat.Rtf
						Console.WriteLine(Constants.vbTab & "RTF format.")
					Case LoadFormat.WordML
						Console.WriteLine(Constants.vbTab & "Microsoft Word 2003 WordprocessingML format.")
					Case LoadFormat.Html
						Console.WriteLine(Constants.vbTab & "HTML format.")
					Case LoadFormat.Mhtml
						Console.WriteLine(Constants.vbTab & "MHTML (Web archive) format.")
					Case LoadFormat.Odt
						Console.WriteLine(Constants.vbTab & "OpenDocument Text.")
					Case LoadFormat.Ott
						Console.WriteLine(Constants.vbTab & "OpenDocument Text Template.")
					Case LoadFormat.DocPreWord97
						Console.WriteLine(Constants.vbTab & "MS Word 6 or Word 95 format.")
					Case Else
						Console.WriteLine(Constants.vbTab & "Unknown format.")
				End Select

				' Now copy the document into the appropriate folder.
				If info.IsEncrypted Then
					Console.WriteLine(Constants.vbTab & "An encrypted document.")
					File.Copy(fileName, Path.Combine(encryptedDir, nameOnly), True)
				Else
					Select Case info.LoadFormat
						Case LoadFormat.DocPreWord97
							File.Copy(fileName, Path.Combine(pre97Dir, nameOnly), True)
						Case LoadFormat.Unknown
							File.Copy(fileName, Path.Combine(unknownDir, nameOnly), True)
						Case Else
							File.Copy(fileName, Path.Combine(supportedDir, nameOnly), True)
					End Select
				End If
			Next fileName
			'ExEnd
		End Sub
	End Class
End Namespace