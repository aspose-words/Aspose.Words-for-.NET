'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////
'14/9/06 by Vladimir Averkin

Imports Microsoft.VisualBasic
Imports System
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Reflection

Namespace Word2Help
	''' <summary>
	''' This project converts documentation stored inside a DOC format to a series of HTML documents. This output is in 
	''' a form that can then be easily compiled together into a single compiled help file (CHM) by using the Microsoft HTML Help workshop application.
	''' </summary>
	Public Class Starter
		Public Shared Sub Main(ByVal args() As String)
			Dim exeDir As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar
			' Specifies the source directory, processes all *.doc files found in it.
			Dim dataDir As String = New Uri(New Uri(exeDir), "../../Data/").LocalPath
			' Specifies the destination directory where the HTML files are output.
			Dim outDir As String = Path.Combine(dataDir, "Out")
			' Specifies the part of the URLs to remove. If there are any hyperlinks that start
			' with the above URL, this URL is removed. This allows the document designer to include
			' links to the HTML API and they will be "corrected" so they work both in the online
			' HTML and also in the compiled CHM.
			Dim fixUrl As String = ""

			' Remove any existing output and recreate the Out folder.
			If Directory.Exists(outDir) Then
				Directory.Delete(outDir, True)
			End If

			Directory.CreateDirectory(outDir)

			''' *** LICENSING ***
			' An Aspose.Words license is required to use this project fully.
			' Without a license Aspose.Words will work in evaluation mode and truncate documents
			' and output watermarks.
			'
			' You can download a free 30-day trial license from the Aspose site. The easiest way is to set the license is to
			' embed it as an embedded resource in this project and uncomment the following code.
			'
			' Aspose.Words.License license = new Aspose.Words.License();
			' license.SetLicense("Aspose.Words.lic");
			Console.WriteLine("Extracting topics from {0}.", dataDir)

			Dim topics As New TopicCollection(dataDir, fixUrl)
			topics.AddFromDir(dataDir)
			topics.WriteHtml(outDir)
			topics.WriteContentXml(outDir)

			Console.WriteLine()
			Console.WriteLine("Conversion completed successfully.")
		End Sub
	End Class
End Namespace
