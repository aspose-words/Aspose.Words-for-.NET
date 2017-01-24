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
	Public Class ExDigitalSignatureCollection
		Inherits ApiExampleBase
		<Test> _
		Public Sub GetEnumeratorEx()
			'ExStart
			'ExFor:DigitalSignatureCollection.GetEnumerator
			'ExSummary:Shows how to load and enumerate all digital signatures of a document.
			Dim digitalSignatures As DigitalSignatureCollection = DigitalSignatureUtil.LoadSignatures(MyDir & "Document.Signed.doc")

			Dim enumerator = digitalSignatures.GetEnumerator()
			Do While enumerator.MoveNext()
				' Do something useful
				Dim ds As DigitalSignature = CType(enumerator.Current, DigitalSignature)
				Console.WriteLine(ds.ToString())
			Loop
			'ExEnd
		End Sub
	End Class
End Namespace