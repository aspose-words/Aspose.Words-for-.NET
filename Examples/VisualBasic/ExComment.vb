'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports Aspose.Words
Imports NUnit.Framework

Namespace Examples
	<TestFixture> _
	Public Class ExComment
		Inherits ExBase
		<Test> _
		Public Sub AcceptAllRevisions()
			'ExStart
			'ExFor:Document.AcceptAllRevisions
			'ExId:AcceptAllRevisions
			'ExSummary:Shows how to accept all tracking changes in the document.
			Dim doc As New Document(MyDir & "Document.doc")
			doc.AcceptAllRevisions()
			'ExEnd
		End Sub
	End Class
End Namespace
