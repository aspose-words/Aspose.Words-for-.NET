'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.IO
Imports System.Reflection

Imports Aspose.Words

Namespace ExtractContentBasedOnStyles
	''' <summary>
	''' Shows how to find paragraphs and runs formatted with a specific style.
	''' </summary>
	Friend Class Program
		Public Shared Sub Main(ByVal args() As String)
			' Sample infrastructure.
			Dim exeDir As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar
			Dim dataDir As String = New Uri(New Uri(exeDir), "../../Data/").LocalPath

			'ExStart
			'ExId:ExtractContentBasedOnStyles_Main
			'ExSummary:Run queries and display results.
			' Open the document.
			Dim doc As New Document(dataDir & "TestFile.doc")

			' Define style names as they are specified in the Word document.
			Const paraStyle As String = "Heading 1"
			Const runStyle As String = "Intense Emphasis"

			' Collect paragraphs with defined styles. 
			' Show the number of collected paragraphs and display the text of this paragraphs.
			Dim paragraphs As ArrayList = ParagraphsByStyleName(doc, paraStyle)
			Console.WriteLine(String.Format("Paragraphs with ""{0}"" styles ({1}):", paraStyle, paragraphs.Count))
			For Each paragraph As Paragraph In paragraphs
				Console.Write(paragraph.ToTxt())
			Next paragraph

			' Collect runs with defined styles. 
			' Show the number of collected runs and display the text of this runs.
			Dim runs As ArrayList = RunsByStyleName(doc, runStyle)
			Console.WriteLine(String.Format(Constants.vbLf & "Runs with ""{0}"" styles ({1}):", runStyle, runs.Count))
			For Each run As Run In runs
				Console.WriteLine(run.Range.Text)
			Next run
			'ExEnd
		End Sub

		'ExStart
		'ExId:ExtractContentBasedOnStyles_Paragraphs
		'ExSummary:Find all paragraphs formatted with the specified style.
		Public Shared Function ParagraphsByStyleName(ByVal doc As Document, ByVal styleName As String) As ArrayList
			' Create an array to collect paragraphs of the specified style.
			Dim paragraphsWithStyle As New ArrayList()
			' Get all paragraphs from the document.
			Dim paragraphs As NodeCollection = doc.GetChildNodes(NodeType.Paragraph, True)
			' Look through all paragraphs to find those with the specified style.
			For Each paragraph As Paragraph In paragraphs
				If paragraph.ParagraphFormat.Style.Name = styleName Then
					paragraphsWithStyle.Add(paragraph)
				End If
			Next paragraph
			Return paragraphsWithStyle
		End Function
		'ExEnd

		'ExStart
		'ExId:ExtractContentBasedOnStyles_Runs
		'ExSummary:Find all runs formatted with the specified style.
		Public Shared Function RunsByStyleName(ByVal doc As Document, ByVal styleName As String) As ArrayList
			' Create an array to collect runs of the specified style.
			Dim runsWithStyle As New ArrayList()
			' Get all runs from the document.
			Dim runs As NodeCollection = doc.GetChildNodes(NodeType.Run, True)
			' Look through all runs to find those with the specified style.
			For Each run As Run In runs
				If run.Font.Style.Name = styleName Then
					runsWithStyle.Add(run)
				End If
			Next run
			Return runsWithStyle
		End Function
		'ExEnd
	End Class
End Namespace