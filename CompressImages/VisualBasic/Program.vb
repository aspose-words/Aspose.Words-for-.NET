'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Reflection
Imports System.Diagnostics

Imports Aspose.Words
Imports Aspose.Words.Drawing

Namespace CompressImages
	Friend Class Program
		Public Shared Sub Main(ByVal args() As String)
			' Sample infrastructure.
			Dim exeDir As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar
			Dim dataDir As String = New Uri(New Uri(exeDir), "../../Data/").LocalPath
			Dim srcFileName As String = dataDir & "Test.docx"

			Console.WriteLine("Loading {0}. Size {1}.", srcFileName, GetFileSize(srcFileName))
			Dim doc As New Document(srcFileName)

			' 220ppi Print - said to be excellent on most printers and screens.
			' 150ppi Screen - said to be good for web pages and projectors.
			' 96ppi Email - said to be good for minimal document size and sharing.
			Const desiredPpi As Integer = 150

			' In .NET this seems to be a good compression / quality setting.
			Const jpegQuality As Integer = 90

			' Resample images to desired ppi and save.
			Dim count As Integer = Resampler.Resample(doc, desiredPpi, jpegQuality)

			Console.WriteLine("Resampled {0} images.", count)

			If count <> 1 Then
				Console.WriteLine("We expected to have only 1 image resampled in this test document!")
			End If

			Dim dstFileName As String = srcFileName & ".Resampled Out.docx"
			doc.Save(dstFileName)
			Console.WriteLine("Saving {0}. Size {1}.", dstFileName, GetFileSize(dstFileName))

			' Verify that the first image was compressed by checking the new Ppi.
			doc = New Document(dstFileName)
			Dim shape As DrawingML = CType(doc.GetChild(NodeType.DrawingML, 0, True), DrawingML)
			Dim imagePpi As Double = shape.ImageData.ImageSize.WidthPixels / ConvertUtil.PointToInch(shape.Size.Width)

			Debug.Assert(imagePpi < 150, "Image was not resampled successfully.")

			Console.WriteLine("Press any key.")
			Console.ReadLine()
		End Sub
		Public Shared Function GetFileSize(ByVal fileName As String) As Integer
			Using stream As Stream = File.OpenRead(fileName)
				Return CInt(Fix(stream.Length))
			End Using
		End Function
	End Class
End Namespace
