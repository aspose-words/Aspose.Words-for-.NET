'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Reflection

Namespace SplitIntoHtmlPagesExample
	Public Class Program
		Public Shared Sub Main()
			' You need to have a valid license for Aspose.Words.
			' The best way is to embed the license as a resource into the project
			' and specify only file name without path in the following call.
			' Aspose.Words.License license = new Aspose.Words.License();
			' license.SetLicense(@"Aspose.Words.lic");


			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			Dim srcFileName As String = dataDir & "SOI 2007-2012-DeeM with footnote added.doc"
			Dim tocTemplate As String = dataDir & "TocTemplate.doc"

			Dim outDir As String = Path.Combine(dataDir, "Out")
			Directory.CreateDirectory(outDir)

			' This class does the job.
			Dim w As New Worker()
			w.Execute(srcFileName, tocTemplate, outDir)

			Console.WriteLine("Success.")
		End Sub
	End Class
End Namespace