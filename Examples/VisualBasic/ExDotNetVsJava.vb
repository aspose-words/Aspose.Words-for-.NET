'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
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
	''' <summary>
	''' Examples for the .NET vs Java Differences in Aspose.Words in the Programmers Guide.
	''' </summary>
	<TestFixture> _
	Public Class ExDotNetVsJava
		'ExStart
		'ExId:SaveSignature
		'ExSummary:Shows difference in .NET and Java in signatures of a method with an enum parameter.
		' The saveFormat parameter is a SaveFormat enum value.
		Private Sub Save(ByVal fileName As String, ByVal saveFormat As SaveFormat)
		'ExEnd
			' Do nothing.
		End Sub

		'ExStart
		'ExId:CollectionItemSignature
		'ExSummary:Shows difference in signatures of collection indexers in .NET vs Java.
		Public Class HeaderFooterCollection
			' Get by index is an indexer.
			Default Public ReadOnly Property Item(ByVal index As Integer) As HeaderFooter
				Get 'ExSkip
					Return Nothing 'ExSkip
				End Get 'ExSkip
			End Property 'ExSkip

			' Get by header footer type is an overloaded indexer.
			Default Public ReadOnly Property Item(ByVal headerFooterType As HeaderFooterType) As HeaderFooter
				Get 'ExSkip
					Return Nothing 'ExSkip
				End Get 'ExSkip
			End Property 'ExSkip
		End Class
		'ExEnd
	End Class
End Namespace
