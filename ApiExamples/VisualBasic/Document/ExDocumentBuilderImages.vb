' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports NUnit.Framework


Namespace ApiExamples.Document
	<TestFixture> _
	Public Class ExDocumentBuilderImages
		Inherits ApiExampleBase
		<Test> _
		Public Sub InsertImageStreamRelativePositionEx()
			'ExStart
			'ExFor:DocumentBuilder.InsertImage(Stream, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
			'ExSummary:Shows how to insert an image into a document from a stream, also using relative positions.
			Dim doc As New Aspose.Words.Document()
			Dim builder As New DocumentBuilder(doc)

			Dim stream As System.IO.Stream = System.IO.File.OpenRead(MyDir & "Aspose.Words.gif")
			Try
				builder.InsertImage(stream, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin, 100, 200, 100, WrapType.Square)
			Finally
				stream.Close()
			End Try

			builder.Document.Save(MyDir & "Image.CreateFromStreamRelativePosition Out.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub InsertImageFromByteArrayEx()
			'ExStart
			'ExFor:DocumentBuilder.InsertImage(Byte[])
			'ExSummary:Shows how to import an image into a document from a byte array.
			Dim doc As New Aspose.Words.Document()
			Dim builder As New DocumentBuilder(doc)

			' Prepare a byte array of an image.
			Dim image As System.Drawing.Image = System.Drawing.Image.FromFile(MyDir & "Aspose.Words.gif")
			Dim imageConverter As New System.Drawing.ImageConverter()
			Dim imageBytes() As Byte = CType(imageConverter.ConvertTo(image, GetType(Byte())), Byte())

			builder.InsertImage(imageBytes)
			builder.Document.Save(MyDir & "Image.CreateFromByteArrayDefault Out.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub InsertImageFromByteArrayCustomSizeEx()
			'ExStart
			'ExFor:DocumentBuilder.InsertImage(Byte[], Double, Double)
			'ExSummary:Shows how to import an image into a document from a byte array, with a custom size.
			Dim doc As New Aspose.Words.Document()
			Dim builder As New DocumentBuilder(doc)

			' Prepare a byte array of an image.
			Dim image As System.Drawing.Image = System.Drawing.Image.FromFile(MyDir & "Aspose.Words.gif")
			Dim imageConverter As New System.Drawing.ImageConverter()
			Dim imageBytes() As Byte = CType(imageConverter.ConvertTo(image, GetType(Byte())), Byte())

			builder.InsertImage(imageBytes, Aspose.Words.ConvertUtil.PixelToPoint(450), Aspose.Words.ConvertUtil.PixelToPoint(144))
			builder.Document.Save(MyDir & "Image.CreateFromByteArrayCustomSize Out.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub InsertImageFromByteArrayRelativePositionEx()
			'ExStart
			'ExFor:DocumentBuilder.InsertImage(Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
			'ExSummary:Shows how to import an image into a document from a byte array, also using relative positions.
			Dim doc As New Aspose.Words.Document()
			Dim builder As New DocumentBuilder(doc)

			' Prepare a byte array of an image.
			Dim image As System.Drawing.Image = System.Drawing.Image.FromFile(MyDir & "Aspose.Words.gif")
			Dim imageConverter As New System.Drawing.ImageConverter()
			Dim imageBytes() As Byte = CType(imageConverter.ConvertTo(image, GetType(Byte())), Byte())

			builder.InsertImage(imageBytes, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin, 100, 200, 100, WrapType.Square)
			builder.Document.Save(MyDir & "Image.CreateFromByteArrayRelativePosition Out.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub InsertImageFromImageCustomSizeEx()
			'ExStart
			'ExFor:DocumentBuilder.InsertImage(Image, Double, Double)
			'ExSummary:Shows how to import an image into a document, with a custom size.
			Dim doc As New Aspose.Words.Document()
			Dim builder As New DocumentBuilder(doc)

			Dim rasterImage As System.Drawing.Image = System.Drawing.Image.FromFile(MyDir & "Aspose.Words.gif")
			Try
				builder.InsertImage(rasterImage, Aspose.Words.ConvertUtil.PixelToPoint(450), Aspose.Words.ConvertUtil.PixelToPoint(144))
				builder.Writeln()
			Finally
				rasterImage.Dispose()
			End Try
			builder.Document.Save(MyDir & "Image.CreateFromImageWithStreamCustomSize Out.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub InsertImageFromImageRelativePositionEx()
			'ExStart
			'ExFor:DocumentBuilder.InsertImage(Image, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
			'ExSummary:Shows how to import an image into a document, also using relative positions.
			Dim doc As New Aspose.Words.Document()
			Dim builder As New DocumentBuilder(doc)

			Dim rasterImage As System.Drawing.Image = System.Drawing.Image.FromFile(MyDir & "Aspose.Words.gif")
			Try
				builder.InsertImage(rasterImage, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin, 100, 200, 100, WrapType.Square)
			Finally
				rasterImage.Dispose()
			End Try

			builder.Document.Save(MyDir & "Image.CreateFromImageWithStreamRelativePosition Out.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub InsertImageStreamCustomSizeEx()
			'ExStart
			'ExFor:DocumentBuilder.InsertImage(Stream, Double, Double)
			'ExSummary:Shows how to import an image from a stream into a document with a custom size.
			Dim doc As New Aspose.Words.Document()
			Dim builder As New DocumentBuilder(doc)

			Dim stream As System.IO.Stream = System.IO.File.OpenRead(MyDir & "Aspose.Words.gif")
			Try
				builder.InsertImage(stream, Aspose.Words.ConvertUtil.PixelToPoint(400), Aspose.Words.ConvertUtil.PixelToPoint(400))
			Finally
				stream.Close()
			End Try

			builder.Document.Save(MyDir & "Image.CreateFromStreamCustomSize Out.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub InsertImageStringCustomSizeEx()
			'ExStart
			'ExFor:DocumentBuilder.InsertImage(String, Double, Double)
			'ExSummary:Shows how to import an image from a url into a document with a custom size.
			Dim doc As New Aspose.Words.Document()
			Dim builder As New DocumentBuilder(doc)

			' Remote URI
			builder.InsertImage("http://www.aspose.com/images/aspose-logo.gif", Aspose.Words.ConvertUtil.PixelToPoint(450), Aspose.Words.ConvertUtil.PixelToPoint(144))

			' Local URI
			builder.InsertImage(MyDir & "Aspose.Words.gif", Aspose.Words.ConvertUtil.PixelToPoint(400), Aspose.Words.ConvertUtil.PixelToPoint(400))

			doc.Save(MyDir & "DocumentBuilder.InsertImageFromUrlCustomSize Out.doc")
			'ExEnd
		End Sub
	End Class
End Namespace
