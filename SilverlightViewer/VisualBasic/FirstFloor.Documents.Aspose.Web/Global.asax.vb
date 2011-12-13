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
Imports Aspose.Words

Namespace FirstFloor.Documents.Aspose.Web
	Public Class [Global]
		Inherits System.Web.HttpApplication

		Protected Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
			' TODO 0 Do not ship source code of this demo project with Aspose.Words.lic embedded in the project. Delete Aspose.Words.lic and this comment before shipping.

			Using licenseStream As Stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("FirstFloor.Documents.Aspose.Web.Aspose.Words.lic")
				If licenseStream IsNot Nothing Then
					Dim license As New License()
					license.SetLicense(licenseStream)
				End If
			End Using
		End Sub
	End Class
End Namespace