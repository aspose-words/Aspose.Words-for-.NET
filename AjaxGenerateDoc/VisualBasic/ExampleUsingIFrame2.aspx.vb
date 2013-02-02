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

Namespace AjaxGenerateDoc
	''' <summary>
	''' Shows how to invoke Aspose.Words for generating a document with data from a GridView control. 
	''' In this example IFrame is used.
	''' </summary>
	Partial Public Class ExampleUsingIFrame2
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
		''' Adds an onClick script that will invoke document generation (in an IFrame).
		''' </summary>
		Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs)
			' We make IFrame to invoke GenerateFile.aspx and pass parameters using a query string.
			Dim script As String = "var iframe = document.createElement('iframe'); " & "iframe.src = 'GenerateFile.aspx?name={0}&company={1}'; " & "iframe.style.display = 'none'; " & "document.body.appendChild(iframe);"

			script = String.Format(script, e.Row.Cells(1).Text, e.Row.Cells(2).Text)

			e.Row.Cells(0).Attributes.Add("onclick", script)
		End Sub
	End Class
End Namespace
