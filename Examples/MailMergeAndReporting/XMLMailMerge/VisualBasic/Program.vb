'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.IO
Imports System.Reflection

Imports Aspose.Words

Namespace XMLMailMergeExample
	Public Class Program
		Public Shared Sub Main()
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create the Dataset and read the XML.
			Dim customersDs As New DataSet()
			customersDs.ReadXml(dataDir & "Customers.xml")

			' Open a template document.
			Dim doc As New Document(dataDir & "TestFile.doc")

			' Execute mail merge to fill the template with data from XML using DataTable.
			doc.MailMerge.Execute(customersDs.Tables("Customer"))

			' Save the output document.
			doc.Save(dataDir & "TestFile Out.doc")
		End Sub
	End Class
End Namespace
'ExEnd