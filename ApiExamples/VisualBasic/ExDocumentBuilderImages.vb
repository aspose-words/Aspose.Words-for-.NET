' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System.Drawing
Imports System.IO

Imports Aspose.Words
Imports Aspose.Words.Drawing

Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Public Class ExDocumentBuilderImages
		Inherits ApiExampleBase
		<Test> _
		Public Sub InsertImageStreamRelativePositionEx()
			'ExStart
			'ExFor:DocumentBuilder.InsertImage(Stream, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
			'ExSummary:Shows how to insert an image into a document from a stream, also using relative positions.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			Dim stream As Stream = File.OpenRead(MyDir & "Aspose.Words.gif")
			Try
				builder.InsertImage(stream, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin, 100, 200, 100, WrapType.Square)
			Finally
				stream.Close()
			End Try

			builder.Document.Save(MyDir & "\Artifacts\Image.CreateFromStreamRelativePosition.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub InsertImageFromByteArrayEx()
			'ExStart
			'ExFor:DocumentBuilder.InsertImage(Byte[])
			'ExSummary:Shows how to import an image into a document from a byte array.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			' Prepare a byte array of an image.
			Dim image As Image = Image.FromFile(MyDir & "Aspose.Words.gif")
			Dim imageConverter As New ImageConverter()
			Dim imageBytes() As Byte = CType(imageConverter.ConvertTo(image, GetType(Byte())), Byte())

			builder.InsertImage(imageBytes)
			builder.Document.Save(MyDir & "\Artifacts\Image.CreateFromByteArrayDefault.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub InsertImageFromByteArrayCustomSizeEx()
			'ExStart
			'ExFor:DocumentBuilder.InsertImage(Byte[], Double, Double)
			'ExSummary:Shows how to import an image into a document from a byte array, with a custom size.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			' Prepare a byte array of an image.
			Dim image As Image = Image.FromFile(MyDir & "Aspose.Words.gif")
			Dim imageConverter As New ImageConverter()
			Dim imageBytes() As Byte = CType(imageConverter.ConvertTo(image, GetType(Byte())), Byte())

			builder.InsertImage(imageBytes, ConvertUtil.PixelToPoint(450), ConvertUtil.PixelToPoint(144))
			builder.Document.Save(MyDir & "\Artifacts\Image.CreateFromByteArrayCustomSize.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub InsertImageFromByteArrayRelativePositionEx()
			'ExStart
			'ExFor:DocumentBuilder.InsertImage(Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
			'ExSummary:Shows how to import an image into a document from a byte array, also using relative positions.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			' Prepare a byte array of an image.
			Dim image As Image = Image.FromFile(MyDir & "Aspose.Words.gif")
			Dim imageConverter As New ImageConverter()
			Dim imageBytes() As Byte = CType(imageConverter.ConvertTo(image, GetType(Byte())), Byte())

			builder.InsertImage(imageBytes, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin, 100, 200, 100, WrapType.Square)
			builder.Document.Save(MyDir & "\Artifacts\Image.CreateFromByteArrayRelativePosition.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub InsertImageFromImageCustomSizeEx()
			'ExStart
			'ExFor:DocumentBuilder.InsertImage(Image, Double, Double)
			'ExSummary:Shows how to import an image into a document, with a custom size.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			Dim rasterImage As Image = Image.FromFile(MyDir & "Aspose.Words.gif")
			Try
				builder.InsertImage(rasterImage, ConvertUtil.PixelToPoint(450), ConvertUtil.PixelToPoint(144))
				builder.Writeln()
			Finally
				rasterImage.Dispose()
			End Try
			builder.Document.Save(MyDir & "\Artifacts\Image.CreateFromImageWithStreamCustomSize.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub InsertImageFromImageRelativePositionEx()
			'ExStart
			'ExFor:DocumentBuilder.InsertImage(Image, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
			'ExSummary:Shows how to import an image into a document, also using relative positions.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			Dim rasterImage As Image = Image.FromFile(MyDir & "Aspose.Words.gif")
			Try
				builder.InsertImage(rasterImage, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin, 100, 200, 100, WrapType.Square)
			Finally
				rasterImage.Dispose()
			End Try

			builder.Document.Save(MyDir & "\Artifacts\Image.CreateFromImageWithStreamRelativePosition.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub InsertImageStreamCustomSizeEx()
			'ExStart
			'ExFor:DocumentBuilder.InsertImage(Stream, Double, Double)
			'ExSummary:Shows how to import an image from a stream into a document with a custom size.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			Dim stream As Stream = File.OpenRead(MyDir & "Aspose.Words.gif")
			Try
				builder.InsertImage(stream, ConvertUtil.PixelToPoint(400), ConvertUtil.PixelToPoint(400))
			Finally
				stream.Close()
			End Try

			builder.Document.Save(MyDir & "\Artifacts\Image.CreateFromStreamCustomSize.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub InsertImageStringCustomSizeEx()
			'ExStart
			'ExFor:DocumentBuilder.InsertImage(String, Double, Double)
			'ExSummary:Shows how to import an image from a url into a document with a custom size.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			' Remote URI
			builder.InsertImage("http://www.aspose.com/images/aspose-logo.gif", ConvertUtil.PixelToPoint(450), ConvertUtil.PixelToPoint(144))

			' Local URI
			builder.InsertImage(MyDir & "Aspose.Words.gif", ConvertUtil.PixelToPoint(400), ConvertUtil.PixelToPoint(400))

			doc.Save(MyDir & "\Artifacts\DocumentBuilder.InsertImageFromUrlCustomSize.doc")
			'ExEnd
		End Sub
	End Class
End Namespace
