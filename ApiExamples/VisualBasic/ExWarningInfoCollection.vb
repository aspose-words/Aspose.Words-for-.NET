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
	Public Class ExWarningInfoCollection
		Inherits ApiExampleBase
		<Test> _
		Public Sub GetEnumeratorEx()
			'ExStart
			'ExFor:WarningInfoCollection.GetEnumerator
			'ExFor:WarningInfoCollection.Clear
			'ExSummary:Shows how to read and clear a collection of warnings.
			Dim wic As New WarningInfoCollection()

			Dim enumerator = wic.GetEnumerator()
			Do While enumerator.MoveNext()
				Dim wi As WarningInfo = CType(enumerator.Current, WarningInfo)
				Console.WriteLine(wi.Description)
			Loop

			wic.Clear()
			'ExEnd
		End Sub
	End Class
End Namespace