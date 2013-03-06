'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.IO

Imports Aspose.Words

Namespace ApplyLicense
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			Dim license As New Aspose.Words.License()

			' This line attempts to set a license from several locations relative to the executable and Aspose.Words.dll.
			' You can also use the additional overload to load a license from a stream, this is useful for instance when the 
			' license is stored as an embedded resource 
			Try
				license.SetLicense("Aspose.Words.lic")

			Catch e As Exception
				' We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license. 
				Console.WriteLine("There was an error setting the license: " & e.Message)
			End Try
		End Sub
	End Class
End Namespace