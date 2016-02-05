' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Collections.Generic
Imports System.Text
Imports NUnit.Framework


Namespace ApiExamples.Hyphenation
	<TestFixture> _
	Public Class ExHyphenation
		Inherits ApiExampleBase
		<Test> _
		Public Sub RegisterDictionaryEx()
			'ExStart
			'ExFor:Hyphenation.RegisterDictionary(String, Stream)
			'ExFor:Hyphenation.RegisterDictionary(String, String)
			'ExSummary:Shows how to open and register a dictionary from a file.
			Dim doc As New Aspose.Words.Document(MyDir & "Document.doc")

			' Register by string
			Aspose.Words.Hyphenation.RegisterDictionary("en-US", MyDir & "hyph_en_US.dic")

			' Register by stream
			Dim dictionaryStream As Stream = New FileStream(MyDir & "hyph_de_CH.dic", FileMode.Open)
			Aspose.Words.Hyphenation.RegisterDictionary("de-CH", dictionaryStream)
			'ExEnd
		End Sub

		<Test> _
		Public Sub IsDictionaryRegisteredEx()
			'ExStart
			'ExFor:Hyphenation.IsDictionaryRegistered(string)
			'ExSummary:Shows how to open check if some dictionary is registered.
			Dim doc As New Aspose.Words.Document(MyDir & "Document.doc")
			Aspose.Words.Hyphenation.RegisterDictionary("en-US", MyDir & "hyph_en_US.dic")

			Console.WriteLine(Aspose.Words.Hyphenation.IsDictionaryRegistered("en-US")) ' True
			'ExEnd
		End Sub

		<Test> _
		Public Sub UnregisterDictionaryEx()
			'ExStart
			'ExFor:Hyphenation.UnregisterDictionary(string)
			'ExSummary:Shows how to un-register a dictionary
			Dim doc As New Aspose.Words.Document(MyDir & "Document.doc")
			Aspose.Words.Hyphenation.RegisterDictionary("en-US", MyDir & "hyph_en_US.dic")

			Aspose.Words.Hyphenation.UnregisterDictionary("en-US")

			Console.WriteLine(Aspose.Words.Hyphenation.IsDictionaryRegistered("en-US")) ' False
			'ExEnd
		End Sub
	End Class
End Namespace
