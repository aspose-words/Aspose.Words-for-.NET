'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Windows.Forms

Imports Aspose.Words

Namespace DocumentInDBExample
	Partial Public Class MainForm
		Inherits Form
		Public Sub New()
			InitializeComponent()

			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")


		End Sub

		<STAThread> _
		Public Shared Sub Main()
			Application.EnableVisualStyles()
			Application.SetCompatibleTextRenderingDefault(False)
			Application.Run(New MainForm())
		End Sub
	End Class
End Namespace