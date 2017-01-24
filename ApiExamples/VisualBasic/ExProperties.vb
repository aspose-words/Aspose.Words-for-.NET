' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System

Imports Aspose.Words
Imports Aspose.Words.Properties

Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Public Class ExProperties
		Inherits ApiExampleBase
		<Test> _
		Public Sub EnumerateProperties()
			'ExStart
			'ExFor:Document.BuiltInDocumentProperties
			'ExFor:Document.CustomDocumentProperties
			'ExFor:BuiltInDocumentProperties
			'ExFor:CustomDocumentProperties
			'ExId:DocumentProperties
			'ExSummary:Enumerates through all built-in and custom properties in a document.
			Dim fileName As String = MyDir & "Properties.doc"
			Dim doc As New Document(fileName)

			Console.WriteLine("1. Document name: {0}", fileName)

			Console.WriteLine("2. Built-in Properties")
			For Each prop As DocumentProperty In doc.BuiltInDocumentProperties
				Console.WriteLine("{0} : {1}", prop.Name, prop.Value)
			Next prop

			Console.WriteLine("3. Custom Properties")
			For Each prop As DocumentProperty In doc.CustomDocumentProperties
				Console.WriteLine("{0} : {1}", prop.Name, prop.Value)
			Next prop
			'ExEnd
		End Sub

		<Test> _
		Public Sub EnumeratePropertiesWithIndexer()
			'ExStart
			'ExFor:DocumentPropertyCollection.Count
			'ExFor:DocumentPropertyCollection.Item(int)
			'ExFor:DocumentProperty
			'ExFor:DocumentProperty.Name
			'ExFor:DocumentProperty.Value
			'ExFor:DocumentProperty.Type
			'ExSummary:Enumerates through all built-in and custom properties in a document using indexed access.
			Dim fileName As String = MyDir & "Properties.doc"
			Dim doc As New Document(fileName)

			Console.WriteLine("1. Document name: {0}", fileName)

			Console.WriteLine("2. Built-in Properties")
			For i As Integer = 0 To doc.BuiltInDocumentProperties.Count - 1
				Dim prop As DocumentProperty = doc.BuiltInDocumentProperties(i)
				Console.WriteLine("{0}({1}) : {2}", prop.Name, prop.Type, prop.Value)
			Next i

			Console.WriteLine("3. Custom Properties")
			For i As Integer = 0 To doc.CustomDocumentProperties.Count - 1
				Dim prop As DocumentProperty = doc.CustomDocumentProperties(i)
				Console.WriteLine("{0}({1}) : {2}", prop.Name, prop.Type, prop.Value)
			Next i
			'ExEnd
		End Sub

		<Test> _
		Public Sub BuiltInNamedAccess()
			'ExStart
			'ExFor:BuiltInDocumentProperties.Item(String)
			'ExFor:DocumentProperty.ToString
			'ExSummary:Retrieves a built-in document property by name.
			Dim doc As New Document(MyDir & "Properties.doc")

			Dim prop As DocumentProperty = doc.BuiltInDocumentProperties("Keywords")
			Console.WriteLine(prop.ToString())
			'ExEnd
		End Sub

		<Test> _
		Public Sub BuiltInPropertiesDirectAccess()
			'ExStart
			'ExFor:BuiltInDocumentProperties.Author
			'ExFor:BuiltInDocumentProperties.Bytes
			'ExFor:BuiltInDocumentProperties.Category
			'ExFor:BuiltInDocumentProperties.Characters
			'ExFor:BuiltInDocumentProperties.CharactersWithSpaces
			'ExFor:BuiltInDocumentProperties.Comments
			'ExFor:BuiltInDocumentProperties.Company
			'ExFor:BuiltInDocumentProperties.CreatedTime
			'ExFor:BuiltInDocumentProperties.Keywords
			'ExFor:BuiltInDocumentProperties.LastPrinted
			'ExFor:BuiltInDocumentProperties.LastSavedBy
			'ExFor:BuiltInDocumentProperties.LastSavedTime
			'ExFor:BuiltInDocumentProperties.Lines
			'ExFor:BuiltInDocumentProperties.Manager
			'ExFor:BuiltInDocumentProperties.NameOfApplication
			'ExFor:BuiltInDocumentProperties.Pages
			'ExFor:BuiltInDocumentProperties.Paragraphs
			'ExFor:BuiltInDocumentProperties.RevisionNumber
			'ExFor:BuiltInDocumentProperties.Subject
			'ExFor:BuiltInDocumentProperties.Template
			'ExFor:BuiltInDocumentProperties.Title
			'ExFor:BuiltInDocumentProperties.TotalEditingTime
			'ExFor:BuiltInDocumentProperties.Version
			'ExFor:BuiltInDocumentProperties.Words
			'ExSummary:Retrieves information from the built-in document properties.
			Dim fileName As String = MyDir & "Properties.doc"
			Dim doc As New Document(fileName)

			Console.WriteLine("Document name: {0}", fileName)
			Console.WriteLine("Document author: {0}", doc.BuiltInDocumentProperties.Author)
			Console.WriteLine("Bytes: {0}", doc.BuiltInDocumentProperties.Bytes)
			Console.WriteLine("Category: {0}", doc.BuiltInDocumentProperties.Category)
			Console.WriteLine("Characters: {0}", doc.BuiltInDocumentProperties.Characters)
			Console.WriteLine("Characters with spaces: {0}", doc.BuiltInDocumentProperties.CharactersWithSpaces)
			Console.WriteLine("Comments: {0}", doc.BuiltInDocumentProperties.Comments)
			Console.WriteLine("Company: {0}", doc.BuiltInDocumentProperties.Company)
			Console.WriteLine("Create time: {0}", doc.BuiltInDocumentProperties.CreatedTime)
			Console.WriteLine("Keywords: {0}", doc.BuiltInDocumentProperties.Keywords)
			Console.WriteLine("Last printed: {0}", doc.BuiltInDocumentProperties.LastPrinted)
			Console.WriteLine("Last saved by: {0}", doc.BuiltInDocumentProperties.LastSavedBy)
			Console.WriteLine("Last saved: {0}", doc.BuiltInDocumentProperties.LastSavedTime)
			Console.WriteLine("Lines: {0}", doc.BuiltInDocumentProperties.Lines)
			Console.WriteLine("Manager: {0}", doc.BuiltInDocumentProperties.Manager)
			Console.WriteLine("Name of application: {0}", doc.BuiltInDocumentProperties.NameOfApplication)
			Console.WriteLine("Pages: {0}", doc.BuiltInDocumentProperties.Pages)
			Console.WriteLine("Paragraphs: {0}", doc.BuiltInDocumentProperties.Paragraphs)
			Console.WriteLine("Revision number: {0}", doc.BuiltInDocumentProperties.RevisionNumber)
			Console.WriteLine("Subject: {0}", doc.BuiltInDocumentProperties.Subject)
			Console.WriteLine("Template: {0}", doc.BuiltInDocumentProperties.Template)
			Console.WriteLine("Title: {0}", doc.BuiltInDocumentProperties.Title)
			Console.WriteLine("Total editing time: {0}", doc.BuiltInDocumentProperties.TotalEditingTime)
			Console.WriteLine("Version: {0}", doc.BuiltInDocumentProperties.Version)
			Console.WriteLine("Words: {0}", doc.BuiltInDocumentProperties.Words)
			'ExEnd
		End Sub

		<Test> _
		Public Sub CustomNamedAccess()
			'ExStart
			'ExFor:DocumentPropertyCollection.Item(String)
			'ExFor:CustomDocumentProperties.Add(String,DateTime)
			'ExFor:DocumentProperty.ToDateTime
			'ExSummary:Retrieves a custom document property by name.
			Dim doc As New Document(MyDir & "Properties.doc")

			Dim prop As DocumentProperty = doc.CustomDocumentProperties("Authorized Date")

			If prop IsNot Nothing Then
				Console.WriteLine(prop.ToDateTime())
			Else
				Console.WriteLine("The document is not authorized. Authorizing...")
				doc.CustomDocumentProperties.Add("AuthorizedDate", DateTime.Now)
			End If
			'ExEnd
		End Sub

		<Test> _
		Public Sub CustomAdd()
			'ExStart
			'ExFor:CustomDocumentProperties.Add(String,String)
			'ExFor:CustomDocumentProperties.Add(String,Boolean)
			'ExFor:CustomDocumentProperties.Add(String,int)
			'ExFor:CustomDocumentProperties.Add(String,DateTime)
			'ExFor:CustomDocumentProperties.Add(String,Double)
			'ExId:AddCustomProperties
			'ExSummary:Checks if a custom property with a given name exists in a document and adds few more custom document properties.
			Dim doc As New Document(MyDir & "Properties.doc")

			Dim props As CustomDocumentProperties = doc.CustomDocumentProperties

			If props("Authorized") Is Nothing Then
				props.Add("Authorized", True)
				props.Add("Authorized By", "John Smith")
				props.Add("Authorized Date", DateTime.Today)
				props.Add("Authorized Revision", doc.BuiltInDocumentProperties.RevisionNumber)
				props.Add("Authorized Amount", 123.45)
			End If
			'ExEnd
		End Sub

		<Test> _
		Public Sub CustomRemove()
			'ExStart
			'ExFor:DocumentPropertyCollection.Remove
			'ExId:RemoveCustomProperties
			'ExSummary:Removes a custom document property.
			Dim doc As New Document(MyDir & "Properties.doc")

			doc.CustomDocumentProperties.Remove("Authorized Date")
			'ExEnd
		End Sub

		<Test> _
		Public Sub PropertyTypes()
			'ExStart
			'ExFor:DocumentProperty.Type
			'ExFor:DocumentProperty.ToBool
			'ExFor:DocumentProperty.ToInt
			'ExFor:DocumentProperty.ToDouble
			'ExFor:DocumentProperty.ToString
			'ExFor:DocumentProperty.ToDateTime
			'ExFor:PropertyType
			'ExSummary:Retrieves the types and values of the custom document properties.
			Dim doc As New Document(MyDir & "Properties.doc")

			For Each prop As DocumentProperty In doc.CustomDocumentProperties
				Console.WriteLine(prop.Name)
				Select Case prop.Type
					Case PropertyType.String
						Console.WriteLine("It's a string value.")
						Console.WriteLine(prop.ToString())
					Case PropertyType.Boolean
						Console.WriteLine("It's a boolean value.")
						Console.WriteLine(prop.ToBool())
					Case PropertyType.Number
						Console.WriteLine("It's an integer value.")
						Console.WriteLine(prop.ToInt())
					Case PropertyType.DateTime
						Console.WriteLine("It's a date time value.")
						Console.WriteLine(prop.ToDateTime())
					Case PropertyType.Double
						Console.WriteLine("It's a double value.")
						Console.WriteLine(prop.ToDouble())
					Case PropertyType.Other
						Console.WriteLine("Other value.")
					Case Else
						Throw New Exception("Unknown property type.")
				End Select
			Next prop
			'ExEnd
		End Sub
	End Class
End Namespace
