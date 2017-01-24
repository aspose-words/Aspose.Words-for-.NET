' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports Aspose.Words
Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Public Class ExDigitalSignature
		Inherits ApiExampleBase
		<Test> _
		Public Sub ToStringEx()
			'ExStart
			'ExFor:DigitalSignature.ToString
			'ExSummary:Shows how to get the string representation of a signature from a signed document.
			Dim doc As New Document(MyDir & "Document.Signed.docx")
			Console.WriteLine(doc.DigitalSignatures(0).ToString())
			'ExEnd
		End Sub
	End Class
End Namespace