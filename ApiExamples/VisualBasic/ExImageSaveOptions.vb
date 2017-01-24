' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports Aspose.Words
Imports Aspose.Words.Saving

Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Friend Class ExImageSaveOptions
		Inherits ApiExampleBase
		<Test> _
		Public Sub UseGdiEmfRenderer()
			'ExStart
			'ExFor:ImageSaveOptions.UseGdiEmfRenderer
			'ExSummary:Shows how to save metafiles directly without using GDI+ to EMF.
			Dim doc As New Document(MyDir & "SaveOptions.MyraidPro.docx")

			Dim saveOptions As New ImageSaveOptions(SaveFormat.Emf)
			saveOptions.UseGdiEmfRenderer = False
			'ExEnd
		End Sub
	End Class
End Namespace
