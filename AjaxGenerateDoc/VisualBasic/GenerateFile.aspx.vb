'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports Aspose.Words

Namespace AjaxGenerateDoc
	''' <summary>
	''' This page is called inside an IFrame to generate a Microsoft Word document.
	''' 
	''' If the caller passes two parameters on the query string "name" and "company",
	''' they will be inserted into the generated document.
	''' </summary>
	Partial Public Class GenerateFile
		Inherits System.Web.UI.Page
		Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
			Dim name As String = String.Empty
			Dim company As String = String.Empty

			If Request.Params("name") IsNot Nothing Then
				name = Request.Params("name")
			End If

			If Request.Params("company") IsNot Nothing Then
				company = Request.Params("company")
			End If

			'Create a new document.
			Dim doc As New Document()

			'Fill the document with custom data.
			Dim builder As New DocumentBuilder(doc)
			If String.IsNullOrEmpty(name) AndAlso String.IsNullOrEmpty(company) Then
				builder.Writeln("Hello World!")
			Else
				builder.Writeln(String.Format("Hello {0} from {1}!", name, company))
			End If

			'This delay is just for a demo! To simulate a delay when building a very complex document.
			System.Threading.Thread.Sleep(2000)

			' Let the caller know we have finished.
			Session("Completed") = True

			'Send the document to the browser.
			doc.Save(Response, "out.doc", ContentDisposition.Attachment, Nothing)

			Response.End()
		End Sub
	End Class
End Namespace
