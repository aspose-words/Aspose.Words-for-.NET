' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports Aspose.Words
Imports Aspose.Words.Markup

Imports NUnit.Framework

Namespace ApiExamples
	''' <summary>
	''' Tests that verify work with structured document tags in the document 
	''' </summary>
	<TestFixture> _
	Friend Class ExStructuredDocumentTag
		Inherits ApiExampleBase
		<Test> _
		Public Sub RepeatingSection()
			Dim doc As New Document(MyDir & "TestRepeatingSection.docx")
			Dim sdts As NodeCollection = doc.GetChildNodes(NodeType.StructuredDocumentTag, True)

			'Assert that the node have sdttype - RepeatingSection and it's not detected as RichText
			Dim sdt As StructuredDocumentTag = CType(sdts(0), StructuredDocumentTag)
			Assert.AreEqual(SdtType.RepeatingSection, sdt.SdtType)

			'Assert that the node have sdttype - RichText 
			sdt = CType(sdts(1), StructuredDocumentTag)
			Assert.AreNotEqual(SdtType.RepeatingSection, sdt.SdtType)
		End Sub
	End Class
End Namespace
