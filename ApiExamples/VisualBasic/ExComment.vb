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
	Public Class ExComment
		Inherits ApiExampleBase
		<Test> _
		Public Sub SetTextEx()
			'ExStart
			'ExFor:Comment.SetText
			'ExSummary:Shows how to add a comment to a document and set it's text.
			Dim doc As New Document(MyDir & "Document.doc")
			Dim builder As New DocumentBuilder(doc)

			Dim comment As New Comment(doc, "John Doe", "J.D.", DateTime.Now)
			builder.CurrentParagraph.AppendChild(comment)
			comment.SetText("My comment.")
			'ExEnd
		End Sub
	End Class
End Namespace
