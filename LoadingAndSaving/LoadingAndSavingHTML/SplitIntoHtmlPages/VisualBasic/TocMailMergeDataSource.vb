'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.Collections
Imports Aspose.Words.Reporting

Namespace SplitIntoHtmlPagesExample
	''' <summary>
	''' A custom data source for Aspose.Words mail merge.
	''' Returns topic objects.
	''' </summary>
	Friend Class TocMailMergeDataSource
		Implements IMailMergeDataSource
		Friend Sub New(ByVal topics As ArrayList)
			mTopics = topics
			' Initialize to BOF.
			mIndex = -1
		End Sub

		Public Function MoveNext() As Boolean Implements IMailMergeDataSource.MoveNext
			If mIndex < mTopics.Count - 1 Then
				mIndex += 1
				Return True
			Else
				' Reached EOF, return false.
				Return False
			End If
		End Function

		Public Function GetValue(ByVal fieldName As String, <System.Runtime.InteropServices.Out()> ByRef fieldValue As Object) As Boolean Implements IMailMergeDataSource.GetValue
			If fieldName = "TocEntry" Then
				' The template document is supposed to have only one field called "TocEntry".
				fieldValue = mTopics(mIndex)
				Return True
			Else
				fieldValue = Nothing
				Return False
			End If
		End Function

		Public ReadOnly Property TableName() As String Implements IMailMergeDataSource.TableName
			' The template document is supposed to have only one merge region called "TOC".
			Get
				Return "TOC"
			End Get
		End Property

		Public Function GetChildDataSource(ByVal tableName As String) As IMailMergeDataSource Implements IMailMergeDataSource.GetChildDataSource
			Return Nothing
		End Function

		Private ReadOnly mTopics As ArrayList
		Private mIndex As Integer
	End Class
End Namespace