'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Drawing

Namespace CompressImages
	Public Class Resampler
		''' <summary>
		''' Resamples all images in the document that are greater than the specified PPI (pixels per inch) to the specified PPI
		''' and converts them to JPEG with the specified quality setting.
		''' </summary>
		''' <param name="doc">The document to process.</param>
		''' <param name="desiredPpi">Desired pixels per inch. 220 high quality. 150 screen quality. 96 email quality.</param>
		''' <param name="jpegQuality">0 - 100% JPEG quality.</param>
		''' <returns></returns>
		Public Shared Function Resample(ByVal doc As Document, ByVal desiredPpi As Integer, ByVal jpegQuality As Integer) As Integer
			Dim count As Integer = 0

			' Convert VML shapes.
			For Each vmlShape As Shape In doc.GetChildNodes(NodeType.Shape, True, False)
				' It is important to use this method to correctly get the picture shape size in points even if the picture is inside a group shape.
				Dim shapeSizeInPoints As SizeF = vmlShape.SizeInPoints

				If ResampleCore(vmlShape.ImageData, shapeSizeInPoints, desiredPpi, jpegQuality) Then
					count += 1
				End If
			Next vmlShape

			' Convert DrawingML shapes.
			For Each dmlShape As DrawingML In doc.GetChildNodes(NodeType.DrawingML, True, False)
				' In MS Word the size of a DrawingML shape is always in points at the moment.
				Dim shapeSizeInPoints As SizeF = dmlShape.Size
				If ResampleCore(dmlShape.ImageData, shapeSizeInPoints, desiredPpi, jpegQuality) Then
					count += 1
				End If
			Next dmlShape

			Return count
		End Function

		''' <summary>
		''' Resamples one VML or DrawingML image
		''' </summary>
		Private Shared Function ResampleCore(ByVal imageData As IImageData, ByVal shapeSizeInPoints As SizeF, ByVal ppi As Integer, ByVal jpegQuality As Integer) As Boolean
			' The are actually several shape types that can have an image (picture, ole object, ole control), let's skip other shapes.
			If imageData Is Nothing Then
				Return False
			End If

			' An image can be stored in the shape or linked from somewhere else. Let's skip images that do not store bytes in the shape.
			Dim originalBytes() As Byte = imageData.ImageBytes
			If originalBytes Is Nothing Then
				Return False
			End If

			' Ignore metafiles, they are vector drawings and we don't want to resample them.
			Dim imageType As ImageType = imageData.ImageType
			If imageType.Equals(ImageType.Wmf) OrElse imageType.Equals(ImageType.Emf) Then
				Return False
			End If

			Try
				Dim shapeWidthInches As Double = ConvertUtil.PointToInch(shapeSizeInPoints.Width)
				Dim shapeHeightInches As Double = ConvertUtil.PointToInch(shapeSizeInPoints.Height)

				' Calculate the current PPI of the image.
				Dim imageSize As ImageSize = imageData.ImageSize
				Dim currentPpiX As Double = imageSize.WidthPixels / shapeWidthInches
				Dim currentPpiY As Double = imageSize.HeightPixels / shapeHeightInches

				Console.Write("Image PpiX:{0}, PpiY:{1}. ", CInt(Fix(currentPpiX)), CInt(Fix(currentPpiY)))

				' Let's resample only if the current PPI is higher than the requested PPI (e.g. we have extra data we can get rid of).
				If (currentPpiX <= ppi) OrElse (currentPpiY <= ppi) Then
					Console.WriteLine("Skipping.")
					Return False
				End If

				Using srcImage As Image = imageData.ToImage()
					' Create a new image of such size that it will hold only the pixels required by the desired ppi.
					Dim dstWidthPixels As Integer = CInt(Fix(shapeWidthInches * ppi))
					Dim dstHeightPixels As Integer = CInt(Fix(shapeHeightInches * ppi))
					Using dstImage As New Bitmap(dstWidthPixels, dstHeightPixels)
						' Drawing the source image to the new image scales it to the new size.
						Using gr As Graphics = Graphics.FromImage(dstImage)
							gr.InterpolationMode = InterpolationMode.HighQualityBicubic
							gr.DrawImage(srcImage, 0, 0, dstWidthPixels, dstHeightPixels)
						End Using

						' Create JPEG encoder parameters with the quality setting.
						Dim encoderInfo As ImageCodecInfo = GetEncoderInfo(ImageFormat.Jpeg)
						Dim encoderParams As New EncoderParameters()
						encoderParams.Param(0) = New EncoderParameter(Encoder.Quality, jpegQuality)

						' Save the image as JPEG to a memory stream.
						Dim dstStream As New MemoryStream()
						dstImage.Save(dstStream, encoderInfo, encoderParams)

						' If the image saved as JPEG is smaller than the original, store it in the shape.
						Console.WriteLine("Original size {0}, new size {1}.", originalBytes.Length, dstStream.Length)
						If dstStream.Length < originalBytes.Length Then
							dstStream.Position = 0
							imageData.SetImage(dstStream)
							Return True
						End If
					End Using
				End Using
			Catch e As Exception
				' Catch an exception, log an error and continue if cannot process one of the images for whatever reason.
				Console.WriteLine("Error processing an image, ignoring. " & e.Message)
			End Try

			Return False
		End Function

		''' <summary>
		''' Gets the codec info for the specified image format. Throws if cannot find.
		''' </summary>
		Private Shared Function GetEncoderInfo(ByVal format As ImageFormat) As ImageCodecInfo
			Dim encoders() As ImageCodecInfo = ImageCodecInfo.GetImageEncoders()

			For i As Integer = 0 To encoders.Length - 1
				If encoders(i).FormatID = format.Guid Then
					Return encoders(i)
				End If
			Next i

			Throw New Exception("Cannot find a codec.")
		End Function
	End Class
End Namespace
