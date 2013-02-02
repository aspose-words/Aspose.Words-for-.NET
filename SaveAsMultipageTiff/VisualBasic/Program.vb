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
Imports System.Reflection

Imports Aspose.Words
Imports Aspose.Words.Saving

Namespace SaveAsMultipageTiff
	Friend Class Program
		Public Shared Sub Main(ByVal args() As String)
			' Sample infrastructure.
			Dim exeDir As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar
			Dim dataDir As String = New Uri(New Uri(exeDir), "../../Data/").LocalPath

			' Open the document.
			Dim doc As New Document(dataDir & "TestFile.doc")

			'ExStart
			'ExId:SaveAsMultipageTiff_save
			'ExSummary:Convert document to TIFF.
			' Save the document as multipage TIFF.
			doc.Save(dataDir & "TestFile Out.tiff")
			'ExEnd

			'ExStart
			'ExId:SaveAsMultipageTiff_SaveWithOptions
			'ExSummary:Convert to TIFF using customized options        
			'Create an ImageSaveOptions object to pass to the Save method
			Dim options As New ImageSaveOptions(SaveFormat.Tiff)
			options.PageIndex = 0
			options.PageCount = 2
			options.TiffCompression = TiffCompression.Ccitt4
			options.Resolution = 160

			doc.Save(dataDir & "TestFileWithOptions Out.tiff", options)
			'ExEnd
		End Sub
	End Class
End Namespace
