' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System.Drawing
Imports Aspose.Words
Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Public Class ExBorderCollection
		Inherits ApiExampleBase
		<Test> _
		Public Sub GetEnumeratorEx()
			'ExStart
			'ExFor:BorderCollection.GetEnumerator
			'ExSummary:Shows how to enumerate all borders in a collection.
			Dim doc As New Document(MyDir & "Document.Borders.doc")
			Dim builder As New DocumentBuilder(doc)
			Dim borders As BorderCollection = builder.ParagraphFormat.Borders

			Dim enumerator = borders.GetEnumerator()
			Do While enumerator.MoveNext()
				' Do something useful.
				Dim b As Border = CType(enumerator.Current, Border)
				b.Color = Color.RoyalBlue
				b.LineStyle = LineStyle.Double
			Loop

			doc.Save(MyDir & "\Artifacts\Document.ChangedColourBorder.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub ClearFormattingEx()
			'ExStart
			'ExFor:BorderCollection.ClearFormatting
			'ExSummary:Shows how to remove all borders from a paragraph at once.
			Dim doc As New Document(MyDir & "Document.Borders.doc")
			Dim builder As New DocumentBuilder(doc)
			Dim borders As BorderCollection = builder.ParagraphFormat.Borders

			borders.ClearFormatting()
			'ExEnd
		End Sub
	End Class
End Namespace