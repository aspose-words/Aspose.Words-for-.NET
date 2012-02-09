'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Reflection
Imports System.Data
Imports System.Diagnostics

Imports Aspose.Words
Imports Aspose.Words.Reporting

Namespace RemoveEmptyRegions
	Friend Class RemoveEmptyRegions
		Public Shared Sub Main(ByVal args() As String)
			' Sample infrastructure.
			Dim exeDir As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar
			Dim dataDir As String = New Uri(New Uri(exeDir), "../../Data/").LocalPath

			'ExStart
			'ExId:RemoveEmptyRegions
			'ExSummary:Shows how to remove unmerged mail merge regions from the document.
			' Open the document.
			Dim doc As New Document(dataDir & "TestFile.doc")

			' Create a dummy data source containing no data.
			Dim data As New DataSet()

			' Set the appropriate mail merge clean up options to remove any unused regions from the document.
			doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedRegions

			' Execute mail merge which will have no effect as there is no data. However the regions found in the document will be removed
			' automatically as they are unused.
			doc.MailMerge.ExecuteWithRegions(data)

			' Save the output document to disk.
			doc.Save(dataDir & "TestFile.RemoveEmptyRegions Out.doc")
			'ExEnd

			Debug.Assert(doc.MailMerge.GetFieldNames().Length = 0, "Error: There are still unused regions remaining in the document")
		End Sub
	End Class
End Namespace