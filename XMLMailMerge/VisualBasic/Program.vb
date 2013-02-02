'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////
'ExStart
'ExId:XMLMailMerge
'ExSummary:Simple Mail Merge from XML using DataSet.

Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.IO
Imports System.Reflection

Imports Aspose.Words

Namespace XMLMailMerge
	Friend Class Program
		Public Shared Sub Main(ByVal args() As String)
			' Sample infrastructure.
			Dim exeDir As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar
			Dim dataDir As String = New Uri(New Uri(exeDir), "../../Data/").LocalPath

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