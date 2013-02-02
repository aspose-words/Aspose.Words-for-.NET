'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Web.UI.WebControls
Imports Aspose.Words

Namespace AjaxGenerateDoc
	''' <summary>
	''' Shows how to invoke Aspose.Words for generating a document with data from a GridView control. 
	''' In this example full post back is used.
	''' </summary>
	Partial Public Class ExampleUsingPostBackTrigger
		Inherits System.Web.UI.Page
		''' <summary>
		''' Fill GridView with data.
		''' </summary>
		Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
			Dim table As New DataTable()
			table.Columns.Add("Name")
			table.Columns.Add("Company")

			Dim row1 As DataRow = table.NewRow()
			row1("Name") = "Alexey"
			row1("Company") = "Aspose"
			table.Rows.Add(row1)

			Dim row2 As DataRow = table.NewRow()
			row2("Name") = "Ravi"
			row2("Company") = "Yolocounty"
			table.Rows.Add(row2)

			GridView1.DataSource = table
			GridView1.DataBind()
		End Sub

		''' <summary>
		''' Generate file when "generate" row command occurs.
		''' </summary>
		Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As GridViewCommandEventArgs)
			If e.CommandName = "generate" Then
				Dim index As Integer = Convert.ToInt32(e.CommandArgument)
				Dim row As GridViewRow = GridView1.Rows(index)
				Dim name As String = row.Cells(1).Text
				Dim company As String = row.Cells(2).Text

				'Create a new document.
				Dim doc As New Document()
				Dim builder As New DocumentBuilder(doc)

				' Fill the document with custom data.
				builder.Writeln(String.Format("Hello {0} from {1}!", name, company))

				'Send created document to a client browser.
				doc.Save(Response, "out.doc", ContentDisposition.Attachment, Nothing)

				Response.End()
			End If
		End Sub
	End Class
End Namespace
