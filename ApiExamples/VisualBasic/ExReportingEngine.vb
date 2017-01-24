' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Data
Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports Aspose.Words.Reporting
Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Public Class ExReportingEngine
		Inherits ApiExampleBase
		Private ReadOnly _image As String = MyDir & "Test_636_852.gif"

		<Test> _
		Public Sub StretchImagefitHeight()
			Dim doc As Document = DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.Image] -fitHeight>>", ShapeType.TextBox)

			Dim imageStream As New ImageStream(New FileStream(Me._image, FileMode.Open, FileAccess.Read))

			BuildReport(doc, imageStream, "src", ReportBuildOptions.None)

			Dim dstStream As New MemoryStream()
			doc.Save(dstStream, SaveFormat.Docx)

			doc = New Document(dstStream)

			Dim shapes As NodeCollection = doc.GetChildNodes(NodeType.Shape, True)

			For Each shape As Shape In shapes
				' Assert that the image is really insert in textbox 
				Assert.IsTrue(shape.ImageData.HasImage)

				'Assert that width is keeped and height is changed
				Assert.AreNotEqual(346.35, shape.Height)
				Assert.AreEqual(431.5, shape.Width)
			Next shape

			dstStream.Dispose()
		End Sub

		<Test> _
		Public Sub StretchImagefitWidth()
			Dim doc As Document = DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.Image] -fitWidth>>", ShapeType.TextBox)

			Dim imageStream As New ImageStream(New FileStream(Me._image, FileMode.Open, FileAccess.Read))

			BuildReport(doc, imageStream, "src", ReportBuildOptions.None)

			Dim dstStream As New MemoryStream()
			doc.Save(dstStream, SaveFormat.Docx)

			doc = New Document(dstStream)

			Dim shapes As NodeCollection = doc.GetChildNodes(NodeType.Shape, True)

			For Each shape As Shape In shapes
				' Assert that the image is really insert in textbox and 
				Assert.IsTrue(shape.ImageData.HasImage)

				'Assert that height is keeped and width is changed
				Assert.AreNotEqual(431.5, shape.Width)
				Assert.AreEqual(346.35, shape.Height)
			Next shape

			dstStream.Dispose()
		End Sub

		<Test> _
		Public Sub StretchImagefitSize()
			Dim doc As Document = DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.Image] -fitSize>>", ShapeType.TextBox)

			Dim imageStream As New ImageStream(New FileStream(Me._image, FileMode.Open, FileAccess.Read))

			BuildReport(doc, imageStream, "src", ReportBuildOptions.None)

			Dim dstStream As New MemoryStream()
			doc.Save(dstStream, SaveFormat.Docx)

			doc = New Document(dstStream)

			Dim shapes As NodeCollection = doc.GetChildNodes(NodeType.Shape, True)

			For Each shape As Shape In shapes
				' Assert that the image is really insert in textbox 
				Assert.IsTrue(shape.ImageData.HasImage)

				'Assert that height is changed and width is changed
				Assert.AreNotEqual(346.35, shape.Height)
				Assert.AreNotEqual(431.5, shape.Width)
			Next shape

			dstStream.Dispose()
		End Sub

		<Test, ExpectedException(GetType(InvalidOperationException))> _
		Public Sub WithoutMissingMembers()
			Dim builder As New DocumentBuilder()

			'Add templete to the document for reporting engine
			DocumentHelper.InsertBuilderText(builder, New String() { "<<[missingObject.First().id]>>", "<<foreach [in missingObject]>><<[id]>><</foreach>>" })

			'Assert that build report failed without "ReportBuildOptions.AllowMissingMembers"
			BuildReport(builder.Document, New DataSet(), "", ReportBuildOptions.None)
		End Sub

		<Test> _
		Public Sub WithMissingMembers()
			Dim builder As New DocumentBuilder()

			'Add templete to the document for reporting engine
			DocumentHelper.InsertBuilderText(builder, New String() { "<<[missingObject.First().id]>>", "<<foreach [in missingObject]>><<[id]>><</foreach>>" })

			BuildReport(builder.Document, New DataSet(), "", ReportBuildOptions.AllowMissingMembers)

			'Assert that build report success with "ReportBuildOptions.AllowMissingMembers"
			Assert.AreEqual(ControlChar.ParagraphBreak + ControlChar.ParagraphBreak + ControlChar.SectionBreak, builder.Document.GetText())
		End Sub

		Private Shared Sub BuildReport(ByVal document As Document, ByVal dataSource As Object, ByVal dataSourceName As String, ByVal reportBuildOptions As ReportBuildOptions)
			Dim engine As New ReportingEngine()
			engine.Options = reportBuildOptions

			engine.BuildReport(document, dataSource, dataSourceName)
		End Sub
	End Class
End Namespace

Public Class ImageStream
	Public Sub New(ByVal stream As Stream)
		Me.Image = stream
	End Sub

	Private privateImage As Stream
	Public Property Image() As Stream
		Get
			Return privateImage
		End Get
		Set(ByVal value As Stream)
			privateImage = value
		End Set
	End Property
End Class