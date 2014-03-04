'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Words

Namespace SimpleMailMergeExample
	Public Class Program
		Public Shared Sub Main()
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			Dim doc As New Document(dataDir & "Template.doc")

			' Fill the fields in the document with user data.
			doc.MailMerge.Execute(New String() { "FullName", "Company", "Address", "Address2", "City" }, New Object() { "James Bond", "MI5 Headquarters", "Milbank", "", "London" })

			' Saves the document to disk.
			doc.Save(dataDir & "MailMerge Result Out.docx")
		End Sub
	End Class
End Namespace