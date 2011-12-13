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

Imports Aspose.Words

Namespace RemoveEmptyRegions
	Friend Class RemoveEmptyRegions
		Public Shared Sub Main(ByVal args() As String)
			' Sample infrastructure.
			Dim exeDir As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar
			Dim dataDir As String = New Uri(New Uri(exeDir), "../../Data/").LocalPath

			'ExStart
			'ExFor:MailMerge.RemoveEmptyRegions
			'ExId:RemoveEmptyRegions
			'ExSummary:Shows how to remove unmerged mail merge regions from the document.
			' Open the document.
			Dim doc As New Document(dataDir & "TestFile.doc")

			' Create a dummy data source containing two empty DataTables which corresponds to the regions in the document.
			Dim data As New DataSet()
			Dim suppliers As New DataTable()
			Dim storeDetails As New DataTable()
			suppliers.TableName = "Suppliers"
			storeDetails.TableName = "StoreDetails"
			data.Tables.Add(suppliers)
			data.Tables.Add(storeDetails)

			' Set the RemoveEmptyRegions to true in order to remove unmerged mail merge regions from the document.
			doc.MailMerge.RemoveEmptyRegions = True

			' Execute mail merge. It will have no effect as there is no data.
			doc.MailMerge.ExecuteWithRegions(data)

			' Save the output document to disk.
			doc.Save(dataDir & "TestFile.RemoveEmptyRegions Out.doc")
			'ExEnd
		End Sub
	End Class
End Namespace