' Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports Aspose.Words
Imports NUnit.Framework



Namespace ApiExamples.EditableRange
	<TestFixture> _
	Friend Class ExEditableRange
		Inherits ApiExampleBase
		<Test> _
		Public Sub RemoveEx()
			'ExStart
			'ExFor:EditableRange.Remove
			'ExSummary:Shows how to remove an editable range from a document.
			Dim doc As New Aspose.Words.Document(MyDir & "Document.doc")
			Dim builder As New DocumentBuilder(doc)

			' Create an EditableRange so we can remove it. Does not have to be well-formed.
			Dim edRange1Start As EditableRangeStart = builder.StartEditableRange()
			Dim editableRange1 As Aspose.Words.EditableRange = edRange1Start.EditableRange
			builder.Writeln("Paragraph inside editable range")
			Dim edRange1End As EditableRangeEnd = builder.EndEditableRange()

			' Remove the range that was just made.
			editableRange1.Remove()
			'ExEnd
		End Sub
	End Class
End Namespace
