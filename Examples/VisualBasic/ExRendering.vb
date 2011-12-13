'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Diagnostics
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Drawing.Text
Imports System.IO
Imports System.Windows.Forms
Imports Aspose.Words
Imports Aspose.Words.Rendering
Imports Aspose.Words.Saving
Imports NUnit.Framework
Imports Aspose.Words.Fonts
Imports System.Collections

Namespace Examples
	<TestFixture> _
	Public Class ExRendering
		Inherits ExBase
		<Test> _
		Public Sub SaveToPdfDefault()
			'ExStart
			'ExFor:Document.Save(String)
			'ExSummary:Converts a whole document to PDF using default options.
			Dim doc As New Document(MyDir & "Rendering.doc")

			doc.Save(MyDir & "Rendering.SaveToPdfDefault Out.pdf")
			'ExEnd
		End Sub

		<Test> _
		Public Sub SaveToPdfWithOutline()
			'ExStart
			'ExFor:Document.Save(String, SaveOptions)
			'ExFor:PdfSaveOptions
			'ExFor:PdfSaveOptions.HeadingsOutlineLevels
			'ExFor:PdfSaveOptions.ExpandedOutlineLevels
			'ExSummary:Converts a whole document to PDF with three levels in the document outline.
			Dim doc As New Document(MyDir & "Rendering.doc")

			Dim options As New PdfSaveOptions()
			options.HeadingsOutlineLevels = 3
			options.ExpandedOutlineLevels = 1

			doc.Save(MyDir & "Rendering.SaveToPdfWithOutline Out.pdf", options)
			'ExEnd
		End Sub

		<Test> _
		Public Sub SaveToPdfStreamOnePage()
			'ExStart
			'ExFor:PdfSaveOptions.PageIndex
			'ExFor:PdfSaveOptions.PageCount
			'ExFor:Document.Save(Stream, SaveOptions)
			'ExSummary:Converts just one page (third page in this example) of the document to PDF.
			Dim doc As New Document(MyDir & "Rendering.doc")

			Using stream As Stream = File.Create(MyDir & "Rendering.SaveToPdfStreamOnePage Out.pdf")
				Dim options As New PdfSaveOptions()
				options.PageIndex = 2
				options.PageCount = 1
				doc.Save(stream, options)
			End Using
			'ExEnd
		End Sub

		<Test> _
		Public Sub SaveToPdfNoCompression()
			'ExStart
			'ExFor:PdfSaveOptions
			'ExFor:PdfSaveOptions.TextCompression
			'ExFor:PdfTextCompression
			'ExSummary:Saves a document to PDF without compression.
			Dim doc As New Document(MyDir & "Rendering.doc")

			Dim options As New PdfSaveOptions()
			options.TextCompression = PdfTextCompression.None

			doc.Save(MyDir & "Rendering.SaveToPdfNoCompression Out.pdf", options)
			'ExEnd
		End Sub

		<Test> _
		Public Sub SaveAsPdf()
			'ExStart
			'ExFor:PdfSaveOptions.PreserveFormFields
			'ExFor:Document.Save(String)
			'ExFor:Document.Save(Stream, SaveFormat)
			'ExFor:Document.Save(String, SaveOptions)
			'ExId:SaveToPdf_NewAPI
			'ExSummary:Shows how to save a document to the PDF format using the Save method and the PdfSaveOptions class.
			' Open the document
			Dim doc As New Document(MyDir & "Rendering.doc")

			' Option 1: Save document to file in the PDF format with default options
			doc.Save(MyDir & "Rendering.PdfDefaultOptions Out.pdf")

			' Option 2: Save the document to stream in the PDF format with default options
			Dim stream As New MemoryStream()
			doc.Save(stream, SaveFormat.Pdf)
			' Rewind the stream position back to the beginning, ready for use
			stream.Seek(0, SeekOrigin.Begin)

			' Option 3: Save document to the PDF format with specified options
			' Render the first page only and preserve form fields as usable controls and not as plain text
			Dim pdfOptions As New PdfSaveOptions()
			pdfOptions.PageIndex = 0
			pdfOptions.PageCount = 1
			pdfOptions.PreserveFormFields = True
			doc.Save(MyDir & "Rendering.PdfCustomOptions Out.pdf", pdfOptions)
			'ExEnd
		End Sub

		<Test> _
		Public Sub SaveAsXps()
			'ExStart
			'ExFor:XpsSaveOptions
			'ExFor:Document.Save(String)
			'ExFor:Document.Save(Stream, SaveFormat)
			'ExFor:Document.Save(String, SaveOptions)
			'ExId:SaveToXps_NewAPI
			'ExSummary:Shows how to save a document to the Xps format using the Save method and the XpsSaveOptions class.
			' Open the document
			Dim doc As New Document(MyDir & "Rendering.doc")
			' Save document to file in the Xps format with default options
			doc.Save(MyDir & "Rendering.XpsDefaultOptions Out.xps")

			' Save document to stream in the Xps format with default options
			Dim docStream As New MemoryStream()
			doc.Save(docStream, SaveFormat.Xps)
			' Rewind the stream position back to the beginning, ready for use
			docStream.Seek(0, SeekOrigin.Begin)

			' Save document to file in the Xps format with specified options
			' Render the first page only
			Dim xpsOptions As New XpsSaveOptions()
			xpsOptions.PageIndex = 0
			xpsOptions.PageCount = 1
			doc.Save(MyDir & "Rendering.XpsCustomOptions Out.xps", xpsOptions)
			'ExEnd
		End Sub

		<Test> _
		Public Sub SaveAsImage()
			'ExStart
			'ExFor:Document.Save(String)
			'ExFor:Document.Save(Stream, SaveFormat)
			'ExFor:Document.Save(String, SaveOptions)
			'ExId:SaveToImage_NewAPI
			'ExSummary:Shows how to save a document to the Jpeg format using the Save method and the ImageSaveOptions class.
			' Open the document
			Dim doc As New Document(MyDir & "Rendering.doc")
			' Save as a Jpeg image file with default options
			doc.Save(MyDir & "Rendering.JpegDefaultOptions Out.jpg")

			' Save document to stream as a Jpeg with default options
			Dim docStream As New MemoryStream()
			doc.Save(docStream, SaveFormat.Jpeg)
			' Rewind the stream position back to the beginning, ready for use
			docStream.Seek(0, SeekOrigin.Begin)

			' Save document to a Jpeg image with specified options.
			' Render the third page only and set the jpeg quality to 80%
			' In this case we need to pass the desired SaveFormat to the ImageSaveOptions constructor 
			' to signal what type of image to save as.
			Dim imageOptions As New ImageSaveOptions(SaveFormat.Jpeg)
			imageOptions.PageIndex = 2
			imageOptions.PageCount = 1
			imageOptions.JpegQuality = 80
			doc.Save(MyDir & "Rendering.JpegCustomOptions Out.jpg", imageOptions)
			'ExEnd
		End Sub

		<Test> _
		Public Sub SaveToTiffDefault()
			'ExStart
			'ExFor:Document.Save(String)
			'ExSummary:Converts a whole document into a multipage TIFF file using default options.
			Dim doc As New Document(MyDir & "Rendering.doc")

			doc.Save(MyDir & "Rendering.SaveToTiffDefault Out.tiff")
			'ExEnd
		End Sub

		<Test> _
		Public Sub SaveToTiffCompression()
			'ExStart
			'ExFor:TiffCompression
			'ExFor:ImageSaveOptions.TiffCompression
			'ExFor:ImageSaveOptions.PageIndex
			'ExFor:ImageSaveOptions.PageCount
			'ExFor:Document.Save(String, SaveOptions)
			'ExSummary:Converts a page of a Word document into a TIFF image and uses the CCITT compression.
			Dim doc As New Document(MyDir & "Rendering.doc")

			Dim options As New ImageSaveOptions(SaveFormat.Tiff)
			options.TiffCompression = TiffCompression.Ccitt3
			options.PageIndex = 0
			options.PageCount = 1

			doc.Save(MyDir & "Rendering.SaveToTiffCompression Out.tif", options)
			'ExEnd
		End Sub

		<Test> _
		Public Sub SaveToImageResolution()
			'ExStart
			'ExFor:ImageSaveOptions
			'ExFor:ImageSaveOptions.Resolution
			'ExSummary:Renders a page of a Word document into a PNG image at a specific resolution.
			Dim doc As New Document(MyDir & "Rendering.doc")

			Dim options As New ImageSaveOptions(SaveFormat.Png)
			options.Resolution = 300
			options.PageCount = 1

			doc.Save(MyDir & "Rendering.SaveToImageResolution Out.png", options)
			'ExEnd
		End Sub

		<Test> _
		Public Sub SaveToEmf()
			'ExStart
			'ExFor:Document.Save(String, SaveOptions)
			'ExSummary:Converts every page of a DOC file into a separate scalable EMF file.
			Dim doc As New Document(MyDir & "Rendering.doc")

			Dim options As New ImageSaveOptions(SaveFormat.Emf)
			options.PageCount = 1

			For i As Integer = 0 To doc.PageCount - 1
				options.PageIndex = i
				doc.Save(MyDir & "Rendering.SaveToEmf." & i.ToString() & " Out.emf", options)
			Next i
			'ExEnd
		End Sub

		<Test> _
		Public Sub SaveToImageJpegQuality()
			'ExStart
			'ExFor:ImageSaveOptions
			'ExFor:ImageSaveOptions.JpegQuality
			'ExSummary:Converts a page of a Word document into JPEG images of different qualities.
			Dim doc As New Document(MyDir & "Rendering.doc")

			Dim options As New ImageSaveOptions(SaveFormat.Jpeg)

			' Try worst quality.
			options.JpegQuality = 0
			doc.Save(MyDir & "Rendering.SaveToImageJpegQuality0 Out.jpeg", options)

			' Try best quality.
			options.JpegQuality = 100
			doc.Save(MyDir & "Rendering.SaveToImageJpegQuality100 Out.jpeg", options)
			'ExEnd
		End Sub

		<Test> _
		Public Sub SaveToImagePaperColor()
			'ExStart
			'ExFor:ImageSaveOptions
			'ExFor:ImageSaveOptions.PaperColor
			'ExSummary:Renders a page of a Word document into an image with transparent or coloured background.
			Dim doc As New Document(MyDir & "Rendering.doc")

			Dim imgOptions As New ImageSaveOptions(SaveFormat.Png)

			imgOptions.PaperColor = Color.Transparent
			doc.Save(MyDir & "Rendering.SaveToImagePaperColorTransparent Out.png", imgOptions)

			imgOptions.PaperColor = Color.LightCoral
			doc.Save(MyDir & "Rendering.SaveToImagePaperColorCoral Out.png", imgOptions)
			'ExEnd
		End Sub

		<Test> _
		Public Sub SaveToImageStream()
			'ExStart
			'ExFor:Document.Save(Stream, SaveFormat)
			'ExSummary:Saves a document page as a BMP image into a stream.
			Dim doc As New Document(MyDir & "Rendering.doc")

			Dim stream As New MemoryStream()
			doc.Save(stream, SaveFormat.Bmp)

			' Rewind the stream and create a .NET image from it.
			stream.Position = 0

			' Read the stream back into an image.
			Dim image As Image = Image.FromStream(stream)
			'ExEnd
		End Sub

		<Test> _
		Public Sub UpdatePageLayout()
			'ExStart
			'ExFor:StyleCollection.Item(String)
			'ExFor:SectionCollection.Item(Int32)
			'ExFor:Document.UpdatePageLayout
			'ExSummary:Shows when to request page layout of the document to be recalculated.
			Dim doc As New Document(MyDir & "Rendering.doc")

			' Saving a document to PDF or to image or printing for the first time will automatically
			' layout document pages and this information will be cached inside the document.
			doc.Save(MyDir & "Rendering.UpdatePageLayout1 Out.pdf")

			' Modify the document in any way.
			doc.Styles("Normal").Font.Size = 6
			doc.Sections(0).PageSetup.Orientation = Aspose.Words.Orientation.Landscape

			' In the current version of Aspose.Words, modifying the document does not automatically rebuild 
			' the cached page layout. If you want to save to PDF or render a modified document again,
			' you need to manually request page layout to be updated.
			doc.UpdatePageLayout()

			doc.Save(MyDir & "Rendering.UpdatePageLayout2 Out.pdf")
			'ExEnd
		End Sub

		<Test> _
		Public Sub UpdateFieldsBeforeRendering()
			'ExStart
			'ExFor:Document.UpdateFields
			'ExId:UpdateFieldsBeforeRendering
			'ExSummary:Shows how to update all fields before rendering a document.
			Dim doc As New Document(MyDir & "Rendering.doc")

			' This updates all fields in the document.
			doc.UpdateFields()

			doc.Save(MyDir & "Rendering.UpdateFields Out.pdf")
			'ExEnd
		End Sub

		<Test, Explicit> _
		Public Sub Print()
			'ExStart
			'ExFor:Document.Print
			'ExSummary:Prints the whole document to the default printer.
			Dim doc As New Document(MyDir & "Document.doc")

			doc.Print()
			'ExEnd
		End Sub

		<Test, Explicit> _
		Public Sub PrintToNamedPrinter()
			'ExStart
			'ExFor:Document.Print(String)
			'ExSummary:Prints the whole document to a specified printer.
			Dim doc As New Document(MyDir & "Document.doc")

			doc.Print("KONICA MINOLTA magicolor 2400W")
			'ExEnd
		End Sub

		<Test, Explicit> _
		Public Sub PrintRange()
			'ExStart
			'ExFor:Document.Print(PrinterSettings)
			'ExSummary:Prints a range of pages.
			Dim doc As New Document(MyDir & "Rendering.doc")

			Dim printerSettings As New PrinterSettings()
			' Page numbers in the .NET printing framework are 1-based.
			printerSettings.FromPage = 1
			printerSettings.ToPage = 3

			doc.Print(printerSettings)
			'ExEnd
		End Sub

		<Test, Explicit> _
		Public Sub PrintRangeWithDocumentName()
			'ExStart
			'ExFor:Document.Print(PrinterSettings, String)
			'ExSummary:Prints a range of pages along with the name of the document.
			Dim doc As New Document(MyDir & "Rendering.doc")

			Dim printerSettings As New PrinterSettings()
			' Page numbers in the .NET printing framework are 1-based.
			printerSettings.FromPage = 1
			printerSettings.ToPage = 3

			doc.Print(printerSettings, "My Print Document.doc")
			'ExEnd
		End Sub

		<Test, Explicit> _
		Public Sub PreviewAndPrint()
			'ExStart
			'ExFor:AsposeWordsPrintDocument
			'ExSummary:Shows the Print dialog that allows selecting the printer and page range to print with. Then brings up the print preview from which you can preview the document and choose to print or close.
			Dim doc As New Document(MyDir & "Rendering.doc")

			Dim previewDlg As New PrintPreviewDialog()
			' Show non-modal first is a hack for the print preview form to show on top.
			previewDlg.Show()

			' Initialize the Print Dialog with the number of pages in the document.
			Dim printDlg As New PrintDialog()
			printDlg.AllowSomePages = True
			printDlg.PrinterSettings.MinimumPage = 1
			printDlg.PrinterSettings.MaximumPage = doc.PageCount
			printDlg.PrinterSettings.FromPage = 1
			printDlg.PrinterSettings.ToPage = doc.PageCount

			If (Not printDlg.ShowDialog().Equals(DialogResult.OK)) Then
				Return
			End If

			' Create the Aspose.Words' implementation of the .NET print document 
			' and pass the printer settings from the dialog to the print document.
			Dim awPrintDoc As New AsposeWordsPrintDocument(doc)
			awPrintDoc.PrinterSettings = printDlg.PrinterSettings

			' Hide and invalidate preview is a hack for print preview to show on top.
			previewDlg.Hide()
			previewDlg.PrintPreviewControl.InvalidatePreview()

			' Pass the Aspose.Words' print document to the .NET Print Preview dialog.
			previewDlg.Document = awPrintDoc

			previewDlg.ShowDialog()
			'ExEnd
		End Sub

		<Test> _
		Public Sub RenderToScale()
			'ExStart
			'ExFor:Document.RenderToScale
			'ExFor:Document.GetPageInfo
			'ExFor:PageInfo
			'ExFor:PageInfo.GetSizeInPixels
			'ExSummary:Renders a page of a Word document into a bitmap using a specified zoom factor.
			Dim doc As New Document(MyDir & "Rendering.doc")

			Dim pageInfo As PageInfo = doc.GetPageInfo(0)

			' Let's say we want the image at 50% zoom.
			Const MyScale As Single = 0.50f

			' Let's say we want the image at this resolution.
			Const MyResolution As Single = 200.0f

			Dim pageSize As Size = pageInfo.GetSizeInPixels(MyScale, MyResolution)

			Using img As New Bitmap(pageSize.Width, pageSize.Height)
				img.SetResolution(MyResolution, MyResolution)

				Using gr As Graphics = Graphics.FromImage(img)
					' You can apply various settings to the Graphics object.
					gr.TextRenderingHint = TextRenderingHint.AntiAliasGridFit

					' Fill the page background.
					gr.FillRectangle(Brushes.White, 0, 0, pageSize.Width, pageSize.Height)

					' Render the page using the zoom.
					doc.RenderToScale(0, gr, 0, 0, MyScale)
				End Using

				img.Save(MyDir & "Rendering.RenderToScale Out.png")
			End Using
			'ExEnd
		End Sub

		<Test> _
		Public Sub RenderToSize()
			'ExStart
			'ExFor:Document.RenderToSize
			'ExSummary:Render to a bitmap at a specified location and size.
			Dim doc As New Document(MyDir & "Rendering.doc")

			Using bmp As New Bitmap(700, 700)
				' User has some sort of a Graphics object. In this case created from a bitmap.
				Using gr As Graphics = Graphics.FromImage(bmp)
					' The user can specify any options on the Graphics object including
					' transform, antialiasing, page units, etc.
					gr.TextRenderingHint = TextRenderingHint.AntiAliasGridFit

					' Let's say we want to fit the page into a 3" x 3" square on the screen so use inches as units.
					gr.PageUnit = GraphicsUnit.Inch

					' The output should be offset 0.5" from the edge and rotated.
					gr.TranslateTransform(0.5f, 0.5f)
					gr.RotateTransform(10)

					' This is our test rectangle.
					gr.DrawRectangle(New Pen(Color.Black, 3f / 72f), 0f, 0f, 3f, 3f)

					' User specifies (in world coordinates) where on the Graphics to render and what size.
					Dim returnedScale As Single = doc.RenderToSize(0, gr, 0f, 0f, 3f, 3f)

					' This is the calculated scale factor to fit 297mm into 3".
					Console.WriteLine("The image was rendered at {0:P0} zoom.", returnedScale)


					' One more example, this time in millimiters.
					gr.PageUnit = GraphicsUnit.Millimeter

					gr.ResetTransform()

					' Move the origin 10mm 
					gr.TranslateTransform(10, 10)

					' Apply both scale transform and page scale for fun.
					gr.ScaleTransform(0.5f, 0.5f)
					gr.PageScale = 2f

					' This is our test rectangle.
					gr.DrawRectangle(New Pen(Color.Black, 1), 90, 10, 50, 100)

					' User specifies (in world coordinates) where on the Graphics to render and what size.
					doc.RenderToSize(1, gr, 90, 10, 50, 100)


					bmp.Save(MyDir & "Rendering.RenderToSize Out.png")
				End Using
			End Using
			'ExEnd
		End Sub

		<Test> _
		Public Sub createThumbnails()
			'ExStart
			'ExFor:Document.RenderToScale
			'ExSummary:Renders individual pages to graphics to create one image with thumbnails of all pages.

			' The user opens or builds a document.
			Dim doc As New Document(MyDir & "Rendering.doc")

			' This defines the number of columns to display the thumbnails in.
			Const thumbColumns As Integer = 2

			' Calculate the required number of rows for thumbnails.
			' We can now get the number of pages in the document.
			Dim remainder As Integer
			Dim thumbRows As Integer = Math.DivRem(doc.PageCount, thumbColumns, remainder)
			If remainder > 0 Then
				thumbRows += 1
			End If

			' Lets say I want thumbnails to be of this zoom.
			Const scale As Single = 0.25f

			' For simplicity lets pretend all pages in the document are of the same size, 
			' so we can use the size of the first page to calculate the size of the thumbnail.
			Dim thumbSize As Size = doc.GetPageInfo(0).GetSizeInPixels(scale, 96)

			' Calculate the size of the image that will contain all the thumbnails.
			Dim imgWidth As Integer = thumbSize.Width * thumbColumns
			Dim imgHeight As Integer = thumbSize.Height * thumbRows

			Using img As New Bitmap(imgWidth, imgHeight)
				' The user has to provides a Graphics object to draw on.
				' The Graphics object can be created from a bitmap, from a metafile, printer or window.
				Using gr As Graphics = Graphics.FromImage(img)
					gr.TextRenderingHint = TextRenderingHint.AntiAliasGridFit

					' Fill the "paper" with white, otherwise it will be transparent.
					gr.FillRectangle(New SolidBrush(Color.White), 0, 0, imgWidth, imgHeight)

					For pageIndex As Integer = 0 To doc.PageCount - 1
						Dim columnIdx As Integer
						Dim rowIdx As Integer = Math.DivRem(pageIndex, thumbColumns, columnIdx)

						' Specify where we want the thumbnail to appear.
						Dim thumbLeft As Single = columnIdx * thumbSize.Width
						Dim thumbTop As Single = rowIdx * thumbSize.Height

						Dim size As SizeF = doc.RenderToScale(pageIndex, gr, thumbLeft, thumbTop, scale)

						' Draw the page rectangle.
						gr.DrawRectangle(Pens.Black, thumbLeft, thumbTop, size.Width, size.Height)
					Next pageIndex

					img.Save(MyDir & "Rendering.Thumbnails Out.png")
				End Using
			End Using
			'ExEnd
		End Sub

		'ExStart
		'ExFor:PageInfo.GetDotNetPaperSize
		'ExFor:PageInfo.Landscape
		'ExSummary:Shows how to implement your own .NET PrintDocument to completely customize printing of Aspose.Words documents.
		<Test, Explicit> _
		Public Sub CustomPrint()
			Dim doc As New Document(MyDir & "Rendering.doc")

			' Create an instance of our own PrintDocument.
			Dim printDoc As New MyPrintDocument(doc)
			' Specify the page range to print.
			printDoc.PrinterSettings.PrintRange = System.Drawing.Printing.PrintRange.SomePages
			printDoc.PrinterSettings.FromPage = 1
			printDoc.PrinterSettings.ToPage = 1

			' Print our document.
			printDoc.Print()
		End Sub

		''' <summary>
		''' The way to print in the .NET Framework is to implement a class derived from PrintDocument.
		''' This class is an example on how to implement custom printing of an Aspose.Words document.
		''' It selects an appropriate paper size, orientation and paper tray when printing.
		''' </summary>
		Public Class MyPrintDocument
			Inherits PrintDocument
			Public Sub New(ByVal document As Document)
				mDocument = document
			End Sub

			''' <summary>
			''' Called before the printing starts. 
			''' </summary>
			Protected Overrides Sub OnBeginPrint(ByVal e As PrintEventArgs)
				MyBase.OnBeginPrint(e)

				' Initialize the range of pages to be printed according to the user selection.
				Select Case PrinterSettings.PrintRange
					Case System.Drawing.Printing.PrintRange.AllPages
						mCurrentPage = 1
						mPageTo = mDocument.PageCount
					Case System.Drawing.Printing.PrintRange.SomePages
						mCurrentPage = PrinterSettings.FromPage
						mPageTo = PrinterSettings.ToPage
					Case Else
						Throw New InvalidOperationException("Unsupported print range.")
				End Select
			End Sub

			''' <summary>
			''' Called before each page is printed. 
			''' </summary>
			Protected Overrides Sub OnQueryPageSettings(ByVal e As QueryPageSettingsEventArgs)
				MyBase.OnQueryPageSettings(e)

				' A single Word document can have multiple sections that specify pages with different sizes, 
				' orientation and paper trays. This code is called by the .NET printing framework before 
				' each page is printed and we get a chance to specify how the page is to be printed.
				Dim pageInfo As PageInfo = mDocument.GetPageInfo(mCurrentPage - 1)
				e.PageSettings.PaperSize = pageInfo.GetDotNetPaperSize(PrinterSettings)
				' MS Word stores the paper source (printer tray) for each section as a printer-specfic value.
				' To obtain the correct tray value you will need to use the RawKindValue returned
				' by .NET for your printer.
				e.PageSettings.PaperSource.RawKind = pageInfo.PaperTray
				e.PageSettings.Landscape = pageInfo.Landscape
			End Sub

			''' <summary>
			''' Called for each page to render it for printing. 
			''' </summary>
			Protected Overrides Sub OnPrintPage(ByVal e As PrintPageEventArgs)
				MyBase.OnPrintPage(e)

				' Aspose.Words rendering engine creates a page that is drawn from the 0,0 of the paper,
				' but there is some hard margin in the printer and the .NET printing framework
				' renders from there. We need to offset by that hard margin.

				' In .NET 1.1 the hard margin is not available programmatically, lets hardcode to about 4mm.
				Dim hardOffsetX As Single = 20
				Dim hardOffsetY As Single = 20

				' This is in .NET 2.0 only. Uncomment when needed.
'                float hardOffsetX = e.PageSettings.HardMarginX;
'                float hardOffsetY = e.PageSettings.HardMarginY;

				Dim pageIndex As Integer = mCurrentPage - 1
				mDocument.RenderToScale(mCurrentPage, e.Graphics, -hardOffsetX, -hardOffsetY, 1.0f)

				mCurrentPage += 1
				e.HasMorePages = (mCurrentPage <= mPageTo)
			End Sub

			Private ReadOnly mDocument As Document
			Private mCurrentPage As Integer
			Private mPageTo As Integer
		End Class
		'ExEnd

		<Test, Explicit> _
		Public Sub WritePageInfo()
			'ExStart
			'ExFor:PageInfo
			'ExFor:PageInfo.PaperSize
			'ExFor:PageInfo.PaperTray
			'ExFor:PageInfo.Landscape
			'ExFor:PageInfo.WidthInPoints
			'ExFor:PageInfo.HeightInPoints
			'ExSummary:Retrieves page size and orientation information for every page in a Word document.
			Dim doc As New Document(MyDir & "Rendering.doc")

			Console.WriteLine("Document ""{0}"" contains {1} pages.", doc.OriginalFileName, doc.PageCount)

			For i As Integer = 0 To doc.PageCount - 1
				Dim pageInfo As PageInfo = doc.GetPageInfo(i)
				Console.WriteLine("Page {0}. PaperSize:{1} ({2:F0}x{3:F0}pt), Orientation:{4}, PaperTray:{5}", i + 1, pageInfo.PaperSize, pageInfo.WidthInPoints, pageInfo.HeightInPoints,If(pageInfo.Landscape, "Landscape", "Portrait"), pageInfo.PaperTray)
			Next i
			'ExEnd
		End Sub

		<Test> _
		Public Sub SetTrueTypeFontsFolder()
			' Store the font folders currently used so we can restore them later. 
			Dim fontFolders() As String = FontSettings.GetFontsFolders()

			'ExStart
			'ExFor:FontSettings
			'ExFor:FontSettings.SetFontsFolder(String, Boolean)
			'ExId:SetFontsFolderCustomFolder
			'ExSummary:Demonstrates how to set the folder Aspose.Words uses to look for TrueType fonts during rendering.
			Dim doc As New Document(MyDir & "Rendering.doc")

			' Set fonts to be scanned for under the specified directory. Do not search within sub-folders.
			FontSettings.SetFontsFolder("C:\MyFonts\", False)

			doc.Save(MyDir & "Rendering.SetFontsFolder Out.pdf")
			'ExEnd

			' Restore the original folders used to search for fonts.
			FontSettings.SetFontsFolders(fontFolders, True)
		End Sub

		<Test> _
		Public Sub SetFontsFoldersMultipleFolders()
			' Store the font folders currently used so we can restore them later. 
			Dim fontFolders() As String = FontSettings.GetFontsFolders()

			'ExStart
			'ExFor:FontSettings
			'ExFor:FontSettings.SetFontsFolders(String[], Boolean)
			'ExId:SetFontsFoldersMultipleFolders
			'ExSummary:Demonstrates how to set Aspose.Words to look in multiple folders for TrueType fonts when rendering.
			Dim doc As New Document(MyDir & "Rendering.doc")

			' Pass true to the second parameter to search within all sub-folders of the specified folders as well.
			FontSettings.SetFontsFolders(New String() {"C:\MyFonts\", "D:\Misc\Fonts\"}, True)

			doc.Save(MyDir & "Rendering.SetFontsFolders Out.pdf")
			'ExEnd

			' Restore the original folders used to search for fonts.
			FontSettings.SetFontsFolders(fontFolders, True)
		End Sub

		<Test> _
		Public Sub SetFontsFoldersSystemAndCustomFolder()
			' Store the font folders currently used so we can restore them later. 
			Dim origFontFolders() As String = FontSettings.GetFontsFolders()

			'ExStart
			'ExFor:FontSettings
			'ExFor:FontSettings.SetFontsFolders(String[], Boolean)
			'ExId:SetFontsFoldersSystemAndCustomFolder
			'ExSummary:Demonstrates how to set Aspose.Words to look for TrueType fonts in system folders and a custom defined folder as well.
			Dim doc As New Document(MyDir & "Rendering.doc")

			' Retrieve the array of environment-dependent font folders that are searched by default. For example this will contain "Windows\Fonts\" on a Windows machines.
			Dim fontFolders As New ArrayList(FontSettings.GetFontsFolders())

			' Add our custom folder to the list.
			fontFolders.Add("C:\MyFonts\")

			' Convert the list into an array and pass it to the FontSettings class.
			FontSettings.SetFontsFolders(CType(fontFolders.ToArray(GetType(String)), String()), True)

			doc.Save(MyDir & "Rendering.SetFontsFolders Out.pdf")
			'ExEnd

			' Verify that folders are set correctly.
			Assert.True(FontSettings.GetFontsFolders()(0).ToLower().Contains("fonts")) ' Regardless of OS the system fonts path should contain "Fonts".
			Assert.AreEqual("C:\MyFonts\", FontSettings.GetFontsFolders()(1))

			' Restore the original folders used to search for fonts.
			FontSettings.SetFontsFolders(origFontFolders, True)
		End Sub

		<Test> _
		Public Sub SetPdfEncryptionPermissions()
			'ExStart
			'ExFor:PdfEncryptionDetails.#ctor
			'ExFor:PdfSaveOptions.EncryptionDetails
			'ExFor:PdfEncryptionDetails.Permissions
			'ExFor:PdfEncryptionAlgorithm
			'ExFor:PdfPermissions
			'ExFor:PdfEncryptionDetails
			'ExSummary:Demonstrates how to set permissions on a PDF document generated by Aspose.Words.
			Dim doc As New Document(MyDir & "Rendering.doc")

			Dim saveOptions As New PdfSaveOptions()

			' Create encryption details and set owner password.
			Dim encryptionDetails As New PdfEncryptionDetails(String.Empty, "password", PdfEncryptionAlgorithm.RC4_128)

			' Start by disallowing all permissions.
			encryptionDetails.Permissions = PdfPermissions.DisallowAll

			' Extend permissions to allow editing or modifying annotations.
			encryptionDetails.Permissions = PdfPermissions.ModifyAnnotations Or PdfPermissions.DocumentAssembly
			saveOptions.EncryptionDetails = encryptionDetails

			' Render the document to PDF format with the specified permissions.
			doc.Save(MyDir & "Rendering.SpecifyPermissions Out.pdf", saveOptions)
			'ExEnd
		End Sub
	End Class
End Namespace
