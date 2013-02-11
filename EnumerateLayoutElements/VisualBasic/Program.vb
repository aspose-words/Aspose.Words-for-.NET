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
Imports System.Drawing

Imports Aspose.Words
Imports Aspose.Words.Layout
Imports Aspose.Words.Rendering

Namespace EnumerateLayoutElements
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			Dim dataDir As String = Path.GetFullPath("../../Data/")

			Dim doc As New Document(dataDir & "TestFile.docx")

			' This creates an enumerator which is used to "walk" the elements of a rendered document.
			Dim it As New LayoutEnumerator(doc)

			' This sample uses the enumerator to write information about each layout element to the console.
			LayoutInfoWriter.Run(it)

			' This sample adds a border around each layout element and saves each page as a JPEG image to the data directory.
			OutlineLayoutEntitiesRenderer.Run(doc, it, dataDir)
		End Sub
	End Class
End Namespace
