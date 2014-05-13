'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Words
Imports System.Data

Namespace MustacheTemplateSyntax
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			Dim ds As New DataSet()

			ds.ReadXml(dataDir & "Orders.xml")

			' Open a template document.
			Dim doc As New Document(dataDir & "ExecuteTemplate.doc")

			doc.MailMerge.UseNonMergeFields = True

			' Execute mail merge to fill the template with data from XML using DataSet.
			doc.MailMerge.ExecuteWithRegions(ds)

			' Save the output document.
			doc.Save(dataDir & "Output.doc")

		End Sub
	End Class
End Namespace