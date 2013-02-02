'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////
'ExStart
'ExFor:MailMerge.Execute(DataRow)
'ExId:MultipleDocsInMailMerge
'ExSummary:Produce multiple documents during mail merge.

Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports System.Reflection

Imports Aspose.Words

Namespace MultipleDocsInMailMerge
	Friend Class Program
		Public Shared Sub Main(ByVal args() As String)
			'Sample infrastructure.
			Dim exeDir As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar
			Dim dataDir As String = New Uri(New Uri(exeDir), "../../Data/").LocalPath

			ProduceMultipleDocuments(dataDir, "TestFile.doc")
		End Sub

		Public Shared Sub ProduceMultipleDocuments(ByVal dataDir As String, ByVal srcDoc As String)
			' Open the database connection.
			Dim connString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dataDir & "Customers.mdb"
			Dim conn As New OleDbConnection(connString)
			conn.Open()
			Try
				' Get data from a database.
				Dim cmd As New OleDbCommand("SELECT * FROM Customers", conn)
				Dim da As New OleDbDataAdapter(cmd)
				Dim data As New DataTable()
				da.Fill(data)

				' Open the template document.
				Dim doc As New Document(dataDir & srcDoc)

				Dim counter As Integer = 1
				' Loop though all records in the data source.
				For Each row As DataRow In data.Rows
					' Clone the template instead of loading it from disk (for speed).
					Dim dstDoc As Document = CType(doc.Clone(True), Document)

					' Execute mail merge.
					dstDoc.MailMerge.Execute(row)

					' Save the document.
					dstDoc.Save(String.Format(dataDir & "TestFile Out {0}.doc", counter))
					counter += 1
				Next row
			Finally
				' Close the database.
				conn.Close()
			End Try
		End Sub
	End Class
End Namespace
'ExEnd