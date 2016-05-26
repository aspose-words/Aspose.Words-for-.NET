' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

#If (Not JAVA) Then
'ExStart
'ExId:ImportForDigitalSignatures
'ExSummary:The import required to use the X509Certificate2 class.

'ExEnd
#End If


Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Net
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Threading
Imports System.Web

Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports Aspose.Words.Fields
Imports Aspose.Words.Properties
Imports Aspose.Words.Rendering
Imports Aspose.Words.Saving
Imports Aspose.Words.Settings
Imports Aspose.Words.Tables
Imports Aspose.Words.Themes

Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Public Class ExDocument
		Inherits ApiExampleBase
		<Test> _
		Public Sub LicenseFromFileNoPath()
			' Copy a license to the bin folder so the example can execute.
			Dim dstFileName As String = Path.Combine(AssemblyDir, "Aspose.Words.lic")
			File.Copy(TestLicenseFileName, dstFileName)

			'ExStart
			'ExFor:License
			'ExFor:License.#ctor
			'ExFor:License.SetLicense(String)
			'ExId:LicenseFromFileNoPath
			'ExSummary:In this example Aspose.Words will attempt to find the license file in the embedded resources or in the assembly folders.
			Dim license As New License()
			license.SetLicense("Aspose.Words.lic")
			'ExEnd

			' Cleanup by removing the license.
			license.SetLicense("")
			File.Delete(dstFileName)
		End Sub

		<Test> _
		Public Sub LicenseFromStream()
			Dim myStream As Stream = File.OpenRead(TestLicenseFileName)
			Try
				'ExStart
				'ExFor:License.SetLicense(Stream)
				'ExId:LicenseFromStream
				'ExSummary:Initializes a license from a stream.
				Dim license As New License()
				license.SetLicense(myStream)
				'ExEnd
			Finally
				myStream.Close()
			End Try
		End Sub

		<Test> _
		Public Sub DocumentCtor()
			'ExStart
			'ExId:DocumentCtor
			'ExSummary:Shows how to create a blank document. Note the blank document contains one section and one paragraph.
			Dim doc As New Document()
			'ExEnd
		End Sub

		<Test> _
		Public Sub OpenFromFile()
			'ExStart
			'ExFor:Document.#ctor(String)
			'ExId:OpenFromFile
			'ExSummary:Opens a document from a file.
			' Open a document. The file is opened read only and only for the duration of the constructor.
			Dim doc As New Document(MyDir & "Document.doc")
			'ExEnd

			'ExStart
			'ExFor:Document.Save(String)
			'ExId:SaveToFile
			'ExSummary:Saves a document to a file.
			doc.Save(MyDir & "\Artifacts\Document.OpenFromFile.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub OpenAndSaveToFile()
			'ExStart
			'ExId:OpenAndSaveToFile
			'ExSummary:Opens a document from a file and saves it to a different format
			Dim doc As New Document(MyDir & "Document.doc")
			doc.Save(MyDir & "\Artifacts\Document.html")
			'ExEnd
		End Sub

		<Test> _
		Public Sub OpenFromStream()
			'ExStart
			'ExFor:Document.#ctor(Stream)
			'ExId:OpenFromStream
			'ExSummary:Opens a document from a stream.
			' Open the stream. Read only access is enough for Aspose.Words to load a document.
			Dim stream As Stream = File.OpenRead(MyDir & "Document.doc")

			' Load the entire document into memory.
			Dim doc As New Document(stream)

			' You can close the stream now, it is no longer needed because the document is in memory.
			stream.Close()

			' ... do something with the document
			'ExEnd

			Assert.AreEqual("Hello World!" & Constants.vbFormFeed, doc.GetText())
		End Sub

		<Test> _
		Public Sub OpenFromStreamWithBaseUri()
			'ExStart
			'ExFor:Document.#ctor(Stream,LoadOptions)
			'ExFor:LoadOptions
			'ExFor:LoadOptions.BaseUri
			'ExId:DocumentCtor_LoadOptions
			'ExSummary:Opens an HTML document with images from a stream using a base URI.

			' We are opening this HTML file:      
			'    <html>
			'    <body>
			'    <p>Simple file.</p>
			'    <p><img src="Aspose.Words.gif" width="80" height="60"></p>
			'    </body>
			'    </html>
			Dim fileName As String = MyDir & "Document.OpenFromStreamWithBaseUri.html"

			' Open the stream.
			Dim stream As Stream = File.OpenRead(fileName)

			' Open the document. Note the Document constructor detects HTML format automatically.
			' Pass the URI of the base folder so any images with relative URIs in the HTML document can be found.
			Dim loadOptions As New LoadOptions()
			loadOptions.BaseUri = MyDir
			Dim doc As New Document(stream, loadOptions)

			' You can close the stream now, it is no longer needed because the document is in memory.
			stream.Close()

			' Save in the DOC format.
			doc.Save(MyDir & "\Artifacts\Document.OpenFromStreamWithBaseUri.doc")
			'ExEnd

			' Lets make sure the image was imported successfully into a Shape node.
			' Get the first shape node in the document.
			Dim shape As Shape = CType(doc.GetChild(NodeType.Shape, 0, True), Shape)

			' Verify some properties of the image.
			Assert.IsTrue(shape.IsImage)
			Assert.IsNotNull(shape.ImageData.ImageBytes)
			Assert.AreEqual(80.0, ConvertUtil.PointToPixel(shape.Width))
			Assert.AreEqual(60.0, ConvertUtil.PointToPixel(shape.Height))
		End Sub

		<Test> _
		Public Sub OpenDocumentFromWeb()
			'ExStart
			'ExFor:Document.#ctor(Stream)
			'ExSummary:Retrieves a document from a URL and saves it to disk in a different format.
			' This is the URL address pointing to where to find the document.
			Dim url As String = "http://www.aspose.com/demos/.net-components/aspose.words/csharp/general/Common/Documents/DinnerInvitationDemo.doc"

			' The easiest way to load our document from the internet is make use of the 
			' System.Net.WebClient class. Create an instance of it and pass the URL
			' to download from.
			Dim webClient As New WebClient()

			' Download the bytes from the location referenced by the URL.
			Dim dataBytes() As Byte = webClient.DownloadData(url)

			' Wrap the bytes representing the document in memory into a MemoryStream object.
			Dim byteStream As New MemoryStream(dataBytes)

			' Load this memory stream into a new Aspose.Words Document.
			' The file format of the passed data is inferred from the content of the bytes itself. 
			' You can load any document format supported by Aspose.Words in the same way.
			Dim doc As New Document(byteStream)

			' Convert the document to any format supported by Aspose.Words.
			doc.Save(MyDir & "\Artifacts\Document.OpenFromWeb.docx")
			'ExEnd
		End Sub

		<Test> _
		Public Sub InsertHtmlFromWebPage()
			'ExStart
			'ExFor:Document.#ctor(Stream, LoadOptions)
			'ExFor:LoadOptions.#ctor(LoadFormat, String, String)
			'ExFor:LoadFormat
			'ExSummary:Shows how to insert the HTML conntents from a web page into a new document.
			' The url of the page to load 
			Dim url As String = "http://www.aspose.com/"

			' Create a WebClient object to easily extract the HTML from the page.
			Dim client As New WebClient()
			Dim pageSource As String = client.DownloadString(url)
			client.Dispose()

			' Get the HTML as bytes for loading into a stream.
			Dim encoding As Encoding = client.Encoding
			Dim pageBytes() As Byte = encoding.GetBytes(pageSource)

			' Load the HTML into a stream.
			Dim stream As New MemoryStream(pageBytes)

			' The baseUri property should be set to ensure any relative img paths are retrieved correctly.
			Dim options As New LoadOptions(Aspose.Words.LoadFormat.Html, "", url)

			' Load the HTML document from stream and pass the LoadOptions object.
			Dim doc As New Document(stream, options)

			' Save the document to disk.
			' The extension of the filename can be changed to save the document into other formats. e.g PDF, DOCX, ODT, RTF.
			doc.Save(MyDir & "\Artifacts\Document.HtmlPageFromWebpage.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub LoadFormat()
			'ExStart
			'ExFor:Document.#ctor(String,LoadOptions)
			'ExFor:LoadFormat
			'ExSummary:Explicitly loads a document as HTML without automatic file format detection.
			Dim loadOptions As New LoadOptions()
			loadOptions.LoadFormat = Aspose.Words.LoadFormat.Html

			Dim doc As New Document(MyDir & "\Artifacts\Document.LoadFormat.html", loadOptions)
			'ExEnd
		End Sub

		<Test> _
		Public Sub LoadFormatForOldDocuments()
			'ExStart
			'ExFor:LoadFormat.DocPreWord60
			'ExSummary: Shows how to open older binary DOC format for Word6.0/Word95 documents
			Dim loadOptions As New LoadOptions()
			loadOptions.LoadFormat = Aspose.Words.LoadFormat.DocPreWord60

			Dim doc As New Document(MyDir & "\Artifacts\Document.PreWord60.doc", loadOptions)
			'ExEnd
		End Sub

		<Test> _
		Public Sub LoadEncryptedFromFile()
			'ExStart
			'ExFor:Document.#ctor(String,LoadOptions)
			'ExFor:LoadOptions
			'ExFor:LoadOptions.#ctor(String)
			'ExId:OpenEncrypted
			'ExSummary:Loads a Microsoft Word document encrypted with a password.
			Dim doc As New Document(MyDir & "Document.LoadEncrypted.doc", New LoadOptions("qwerty"))
			'ExEnd
		End Sub

		<Test> _
		Public Sub LoadEncryptedFromStream()
			'ExStart
			'ExFor:Document.#ctor(Stream,LoadOptions)
			'ExSummary:Loads a Microsoft Word document encrypted with a password from a stream.
			Dim stream As Stream = File.OpenRead(MyDir & "Document.LoadEncrypted.doc")
			Dim doc As New Document(stream, New LoadOptions("qwerty"))
			stream.Close()
			'ExEnd
		End Sub

		<Test> _
		Public Sub ConvertToHtml()
			'ExStart
			'ExFor:Document.Save(String,SaveFormat)
			'ExFor:SaveFormat
			'ExSummary:Converts from DOC to HTML format.
			Dim doc As New Document(MyDir & "Document.doc")

			doc.Save(MyDir & "\Artifacts\Document.ConvertToHtml.html", SaveFormat.Html)
			'ExEnd
		End Sub

		<Test> _
		Public Sub ConvertToMhtml()
			'ExStart
			'ExFor:Document.Save(String)
			'ExSummary:Converts from DOC to MHTML format.
			Dim doc As New Document(MyDir & "Document.doc")

			doc.Save(MyDir & "\Artifacts\Document.ConvertToMhtml.mht")
			'ExEnd
		End Sub

		<Test> _
		Public Sub ConvertToTxt()
			'ExStart
			'ExId:ExtractContentSaveAsText
			'ExSummary:Shows how to save a document in TXT format.
			Dim doc As New Document(MyDir & "Document.doc")

			doc.Save(MyDir & "\Artifacts\Document.ConvertToTxt.txt")
			'ExEnd
		End Sub

		<Test> _
		Public Sub Doc2PdfSave()
			'ExStart
			'ExFor:Document
			'ExFor:Document.Save(String)
			'ExId:Doc2PdfSave
			'ExSummary:Converts a whole document from DOC to PDF using default options.
			Dim doc As New Document(MyDir & "Document.doc")

			doc.Save(MyDir & "\Artifacts\Document.Doc2PdfSave.pdf")
			'ExEnd
		End Sub

		<Test> _
		Public Sub SaveToStream()
			'ExStart
			'ExFor:Document.Save(Stream,SaveFormat)
			'ExId:SaveToStream
			'ExSummary:Shows how to save a document to a stream.
			Dim doc As New Document(MyDir & "Document.doc")

			Dim dstStream As New MemoryStream()
			doc.Save(dstStream, SaveFormat.Docx)

			' Rewind the stream position back to zero so it is ready for next reader.
			dstStream.Position = 0
			'ExEnd
		End Sub

		''' <summary>
		''' RK We are not actually executing this as a test because it does not seem to work without ASP.NET
		''' </summary>
		Public Sub SaveToBrowser()
			' Create a dummy HTTP response.
			Dim Response As New HttpResponse(Nothing)

			'ExStart
			'ExId:SaveToBrowser
			'ExSummary:Shows how to send a document to the client browser from an ASP.NET code.
			Dim doc As New Document(MyDir & "Document.doc")

			doc.Save(Response, "\Artifacts\Report.doc", ContentDisposition.Inline, Nothing)
			'ExEnd
		End Sub

		<Test> _
		Public Sub Doc2EpubSave()
			'ExStart
			'ExId:Doc2EpubSave
			'ExSummary:Converts a document to EPUB using default save options.

			' Open an existing document from disk.
			Dim doc As New Document(MyDir & "Document.EpubConversion.doc")

			' Save the document in EPUB format.
			doc.Save(MyDir & "\Artifacts\Document.EpubConversion.epub")
			'ExEnd
		End Sub

		<Test> _
		Public Sub Doc2EpubSaveWithOptions()
			'ExStart
			'ExFor:HtmlSaveOptions
			'ExFor:HtmlSaveOptions.#ctor
			'ExFor:HtmlSaveOptions.Encoding
			'ExFor:HtmlSaveOptions.DocumentSplitCriteria
			'ExFor:HtmlSaveOptions.ExportDocumentProperties
			'ExFor:HtmlSaveOptions.SaveFormat
			'ExId:Doc2EpubSaveWithOptions
			'ExSummary:Converts a document to EPUB with save options specified.
			' Open an existing document from disk.
			Dim doc As New Document(MyDir & "Document.EpubConversion.doc")

			' Create a new instance of HtmlSaveOptions. This object allows us to set options that control
			' how the output document is saved.
			Dim saveOptions As New HtmlSaveOptions()

			' Specify the desired encoding.
			saveOptions.Encoding = Encoding.UTF8

			' Specify at what elements to split the internal HTML at. This creates a new HTML within the EPUB 
			' which allows you to limit the size of each HTML part. This is useful for readers which cannot read 
			' HTML files greater than a certain size e.g 300kb.
			saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph

			' Specify that we want to export document properties.
			saveOptions.ExportDocumentProperties = True

			' Specify that we want to save in EPUB format.
			saveOptions.SaveFormat = SaveFormat.Epub

			' Export the document as an EPUB file.
			doc.Save(MyDir & "\Artifacts\Document.EpubConversion.epub", saveOptions)
			'ExEnd
		End Sub

		<Test> _
		Public Sub SaveHtmlPrettyFormat()
			'ExStart
			'ExFor:SaveOptions.PrettyFormat
			'ExSummary:Shows how to pass an option to export HTML tags in a well spaced, human readable format.
			Dim doc As New Document(MyDir & "Document.doc")

			Dim htmlOptions As New HtmlSaveOptions(SaveFormat.Html)
			' Enabling the PrettyFormat setting will export HTML in an indented format that is easy to read.
			' If this is setting is false (by default) then the HTML tags will be exported in condensed form with no indentation.
			htmlOptions.PrettyFormat = True

			doc.Save(MyDir & "\Artifacts\Document.PrettyFormat.html", htmlOptions)
			'ExEnd
		End Sub

		<Test> _
		Public Sub SaveHtmlWithOptions()
			'ExStart
			'ExFor:HtmlSaveOptions
			'ExFor:HtmlSaveOptions.ExportTextInputFormFieldAsText
			'ExFor:HtmlSaveOptions.ImagesFolder
			'ExId:SaveWithOptions
			'ExSummary:Shows how to set save options before saving a document to HTML.
			Dim doc As New Document(MyDir & "Rendering.doc")

			' This is the directory we want the exported images to be saved to.
			Dim imagesDir As String = Path.Combine(MyDir, "Images")

			' The folder specified needs to exist and should be empty.
			If Directory.Exists(imagesDir) Then
				Directory.Delete(imagesDir, True)
			End If

			Directory.CreateDirectory(imagesDir)

			' Set an option to export form fields as plain text, not as HTML input elements.
			Dim options As New HtmlSaveOptions(SaveFormat.Html)
			options.ExportTextInputFormFieldAsText = True
			options.ImagesFolder = imagesDir

			doc.Save(MyDir & "\Artifacts\Document.SaveWithOptions.html", options)
			'ExEnd

			' Verify the images were saved to the correct location.
			Assert.IsTrue(File.Exists(MyDir & "\Artifacts\Document.SaveWithOptions.html"))
			Assert.AreEqual(9, Directory.GetFiles(imagesDir).Length)
		End Sub

		''' <summary>
		''' This calls the below method to resolve skipping of [Test] in VB.NET.
		''' </summary>
		<Test> _
		Public Sub SaveHtmlExportFontsCaller()
			Me.SaveHtmlExportFonts()
		End Sub

		'ExStart
		'ExFor:HtmlSaveOptions.ExportFontResources
		'ExFor:HtmlSaveOptions.FontSavingCallback
		'ExFor:IFontSavingCallback
		'ExFor:IFontSavingCallback.FontSaving
		'ExFor:FontSavingArgs
		'ExFor:FontSavingArgs.FontFamilyName
		'ExFor:FontSavingArgs.FontFileName
		'ExId:SaveHtmlExportFonts
		'ExSummary:Shows how to define custom logic for handling font exporting when saving to HTML based formats.
		Public Sub SaveHtmlExportFonts()
			Dim doc As New Document(MyDir & "Document.doc")

			' Set the option to export font resources.
			Dim options As New HtmlSaveOptions(SaveFormat.Mhtml)
			options.ExportFontResources = True
			' Create and pass the object which implements the handler methods.
			options.FontSavingCallback = New HandleFontSaving()

			doc.Save(MyDir & "\Artifacts\Document.SaveWithFontsExport.html", options)
		End Sub

		Public Class HandleFontSaving
			Implements IFontSavingCallback
			Private Sub IFontSavingCallback_FontSaving(ByVal args As FontSavingArgs) Implements IFontSavingCallback.FontSaving
				' You can implement logic here to rename fonts, save to file etc. For this example just print some details about the current font being handled.
				Console.WriteLine("Font Name = {0}, Font Filename = {1}", args.FontFamilyName, args.FontFileName)
			End Sub
		End Class
		'ExEnd

		''' <summary>
		''' This calls the below method to resolve skipping of [Test] in VB.NET.
		''' </summary>
		<Test> _
		Public Sub SaveHtmlExportImagesCaller()
			Me.SaveHtmlExportImages()
		End Sub

		'ExStart
		'ExFor:IImageSavingCallback
		'ExFor:IImageSavingCallback.ImageSaving
		'ExFor:ImageSavingArgs
		'ExFor:ImageSavingArgs.ImageFileName
		'ExFor:HtmlSaveOptions
		'ExFor:HtmlSaveOptions.ImageSavingCallback
		'ExId:SaveHtmlCustomExport
		'ExSummary:Shows how to define custom logic for controlling how images are saved when exporting to HTML based formats.
		Public Sub SaveHtmlExportImages()
			Dim doc As New Document(MyDir & "Document.doc")

			' Create and pass the object which implements the handler methods.
			Dim options As New HtmlSaveOptions(SaveFormat.Html)
			options.ImageSavingCallback = New HandleImageSaving()

			doc.Save(MyDir & "\Artifacts\Document.SaveWithCustomImagesExport.html", options)
		End Sub

		Public Class HandleImageSaving
			Implements IImageSavingCallback
			Private Sub IImageSavingCallback_ImageSaving(ByVal e As ImageSavingArgs) Implements IImageSavingCallback.ImageSaving
				' Change any images in the document being exported with the extension of "jpeg" to "jpg".
				If e.ImageFileName.EndsWith(".jpeg") Then
					e.ImageFileName = e.ImageFileName.Replace(".jpeg", ".jpg")
				End If
			End Sub
		End Class
		'ExEnd

		''' <summary>
		''' This calls the below method to resolve skipping of [Test] in VB.NET.
		''' </summary>
		<Test> _
		Public Sub TestNodeChangingInDocumentCaller()
			Me.TestNodeChangingInDocument()
		End Sub

		'ExStart
		'ExFor:INodeChangingCallback
		'ExFor:INodeChangingCallback.NodeInserting
		'ExFor:INodeChangingCallback.NodeInserted
		'ExFor:INodeChangingCallback.NodeRemoving
		'ExFor:INodeChangingCallback.NodeRemoved
		'ExFor:NodeChangingArgs
		'ExFor:NodeChangingArgs.Node
		'ExFor:DocumentBase.NodeChangingCallback
		'ExId:NodeChangingInDocument
		'ExSummary:Shows how to implement custom logic over node insertion in the document by changing the font of inserted HTML content.
		Public Sub TestNodeChangingInDocument()
			' Create a blank document object
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			' Set up and pass the object which implements the handler methods.
			doc.NodeChangingCallback = New HandleNodeChangingFontChanger()

			' Insert sample HTML content
			builder.InsertHtml("<p>Hello World</p>")

			doc.Save(MyDir & "\Artifacts\Document.FontChanger.doc")

			' Check that the inserted content has the correct formatting
			Dim run As Run = CType(doc.GetChild(NodeType.Run, 0, True), Run)
			Assert.AreEqual(24.0, run.Font.Size)
			Assert.AreEqual("Arial", run.Font.Name)
		End Sub

		Public Class HandleNodeChangingFontChanger
			Implements INodeChangingCallback
			' Implement the NodeInserted handler to set default font settings for every Run node inserted into the Document
			Private Sub NodeInserted(ByVal args As NodeChangingArgs) Implements INodeChangingCallback.NodeInserted
				' Change the font of inserted text contained in the Run nodes.
				If args.Node.NodeType = NodeType.Run Then
					Dim font As Aspose.Words.Font = (CType(args.Node, Run)).Font
					font.Size = 24
					font.Name = "Arial"
				End If
			End Sub

			Private Sub NodeInserting(ByVal args As NodeChangingArgs) Implements INodeChangingCallback.NodeInserting
				' Do Nothing
			End Sub

			Private Sub NodeRemoved(ByVal args As NodeChangingArgs) Implements INodeChangingCallback.NodeRemoved
				' Do Nothing
			End Sub

			Private Sub NodeRemoving(ByVal args As NodeChangingArgs) Implements INodeChangingCallback.NodeRemoving
				' Do Nothing
			End Sub
		End Class
		'ExEnd

		<Test> _
		Public Sub DetectFileFormat()
			'ExStart
			'ExFor:FileFormatUtil.DetectFileFormat(String)
			'ExFor:FileFormatInfo
			'ExFor:FileFormatInfo.LoadFormat
			'ExFor:FileFormatInfo.IsEncrypted
			'ExFor:FileFormatInfo.HasDigitalSignature
			'ExId:DetectFileFormat
			'ExSummary:Shows how to use the FileFormatUtil class to detect the document format and other features of the document.
			Dim info As FileFormatInfo = FileFormatUtil.DetectFileFormat(MyDir & "Document.doc")
			Console.WriteLine("The document format is: " & FileFormatUtil.LoadFormatToExtension(info.LoadFormat))
			Console.WriteLine("Document is encrypted: " & info.IsEncrypted)
			Console.WriteLine("Document has a digital signature: " & info.HasDigitalSignature)
			'ExEnd
		End Sub

		<Test> _
		Public Sub DetectFileFormat_EnumConversions()
			'ExStart
			'ExFor:FileFormatUtil.DetectFileFormat(Stream)
			'ExFor:FileFormatUtil.LoadFormatToExtension(LoadFormat)
			'ExFor:FileFormatUtil.ExtensionToSaveFormat(String)
			'ExFor:FileFormatUtil.SaveFormatToExtension(SaveFormat)
			'ExFor:FileFormatUtil.LoadFormatToSaveFormat(LoadFormat)
			'ExFor:Document.OriginalFileName
			'ExFor:FileFormatInfo.LoadFormat
			'ExSummary:Shows how to use the FileFormatUtil methods to detect the format of a document without any extension and save it with the correct file extension.
			' Load the document without a file extension into a stream and use the DetectFileFormat method to detect it's format. These are both times where you might need extract the file format as it's not visible
			Dim docStream As FileStream = File.OpenRead(MyDir & "Document.FileWithoutExtension") ' The file format of this document is actually ".doc"
			Dim info As FileFormatInfo = FileFormatUtil.DetectFileFormat(docStream)

			' Retrieve the LoadFormat of the document.
			Dim loadFormat As LoadFormat = info.LoadFormat

			' Let's show the different methods of converting LoadFormat enumerations to SaveFormat enumerations.
			'
			' Method #1
			' Convert the LoadFormat to a string first for working with. The string will include the leading dot in front of the extension.
			Dim fileExtension As String = FileFormatUtil.LoadFormatToExtension(loadFormat)
			' Now convert this extension into the corresponding SaveFormat enumeration
			Dim saveFormat As SaveFormat = FileFormatUtil.ExtensionToSaveFormat(fileExtension)

			' Method #2
			' Convert the LoadFormat enumeration directly to the SaveFormat enumeration.
			saveFormat = FileFormatUtil.LoadFormatToSaveFormat(loadFormat)

			' Load a document from the stream.
			Dim doc As New Document(docStream)

			' Save the document with the original file name, " Out" and the document's file extension.
			doc.Save(MyDir & "\Artifacts\Document.WithFileExtension" & FileFormatUtil.SaveFormatToExtension(saveFormat))
			'ExEnd

			Assert.AreEqual(".doc", FileFormatUtil.SaveFormatToExtension(saveFormat))
		End Sub

		<Test> _
		Public Sub DetectFileFormat_SaveFormatToLoadFormat()
			'ExStart
			'ExFor:FileFormatUtil.SaveFormatToLoadFormat(SaveFormat)
			'ExSummary:Shows how to use the FileFormatUtil class and to convert a SaveFormat enumeration into the corresponding LoadFormat enumeration.
			' Define the SaveFormat enumeration to convert.
			Dim saveFormat As SaveFormat = SaveFormat.Html
			' Convert the SaveFormat enumeration to LoadFormat enumeration.
			Dim loadFormat As LoadFormat = FileFormatUtil.SaveFormatToLoadFormat(saveFormat)
			Console.WriteLine("The converted LoadFormat is: " & FileFormatUtil.LoadFormatToExtension(loadFormat))
			'ExEnd

			Assert.AreEqual(".html", FileFormatUtil.SaveFormatToExtension(saveFormat))
			Assert.AreEqual(".html", FileFormatUtil.LoadFormatToExtension(loadFormat))
		End Sub

		<Test> _
		Public Sub AppendDocument()
			'ExStart
			'ExFor:Document.AppendDocument(Document, ImportFormatMode)
			'ExSummary:Shows how to append a document to the end of another document.
			' The document that the content will be appended to.
			Dim dstDoc As New Document(MyDir & "Document.doc")
			' The document to append.
			Dim srcDoc As New Document(MyDir & "DocumentBuilder.doc")

			' Append the source document to the destination document.
			' Pass format mode to retain the original formatting of the source document when importing it.
			dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting)

			' Save the document.
			dstDoc.Save(MyDir & "\Artifacts\Document.AppendDocument.doc")
			'ExEnd
		End Sub

		' Using this file path keeps the example making sense when compared with automation so we expect
		' the file not to be found.
		<Test, ExpectedException(GetType(FileNotFoundException))> _
		Public Sub AppendDocumentFromAutomation()
			'ExStart
			'ExId:AppendDocumentFromAutomation
			'ExSummary:Shows how to join multiple documents together.
			' The document that the other documents will be appended to.
			Dim doc As New Document()
			' We should call this method to clear this document of any existing content.
			doc.RemoveAllChildren()

			Dim recordCount As Integer = 5
			For i As Integer = 1 To recordCount
				' Open the document to join.
				Dim srcDoc As New Document("C:\DetailsList.doc")

				' Append the source document at the end of the destination document.
				doc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles)

				' In automation you were required to insert a new section break at this point, however in Aspose.Words we 
				' don't need to do anything here as the appended document is imported as separate sectons already.

				' If this is the second document or above being appended then unlink all headers footers in this section 
				' from the headers and footers of the previous section.
				If i > 1 Then
					doc.Sections(i).HeadersFooters.LinkToPrevious(False)
				End If
			Next i
			'ExEnd
		End Sub

		<Test> _
		Public Sub DetectDocumentSignatures()
			'ExStart
			'ExFor:FileFormatUtil.DetectFileFormat(String)
			'ExFor:FileFormatInfo.HasDigitalSignature
			'ExId:DetectDocumentSignatures
			'ExSummary:Shows how to check a document for digital signatures before loading it into a Document object.
			' The path to the document which is to be processed.
			Dim filePath As String = MyDir & "Document.Signed.docx"

			Dim info As FileFormatInfo = FileFormatUtil.DetectFileFormat(filePath)
			If info.HasDigitalSignature Then
				Console.WriteLine("Document {0} has digital signatures, they will be lost if you open/save this document with Aspose.Words.", Path.GetFileName(filePath))
			End If
			'ExEnd
		End Sub

		<Test> _
		Public Sub ValidateAllDocumentSignatures()
			'ExStart
			'ExFor:Document.DigitalSignatures
			'ExFor:DigitalSignatureCollection
			'ExFor:DigitalSignatureCollection.IsValid
			'ExId:ValidateAllDocumentSignatures
			'ExSummary:Shows how to validate all signatures in a document.
			' Load the signed document.
			Dim doc As New Document(MyDir & "Document.Signed.docx")

			If doc.DigitalSignatures.IsValid Then
				Console.WriteLine("Signatures belonging to this document are valid")
			Else
				Console.WriteLine("Signatures belonging to this document are NOT valid")
			End If
			'ExEnd

			Assert.True(doc.DigitalSignatures.IsValid)
		End Sub

		<Test> _
		Public Sub ValidateIndividualDocumentSignatures()
			'ExStart
			'ExFor:DigitalSignature
			'ExFor:Document.DigitalSignatures
			'ExFor:DigitalSignature.IsValid
			'ExFor:DigitalSignature.Comments
			'ExFor:DigitalSignature.SignTime
			'ExFor:DigitalSignature.SignatureType
			'ExFor:DigitalSignature.Certificate
			'ExId:ValidateIndividualSignatures
			'ExSummary:Shows how to validate each signature in a document and display basic information about the signature.
			' Load the document which contains signature.
			Dim doc As New Document(MyDir & "Document.Signed.docx")

			For Each signature As DigitalSignature In doc.DigitalSignatures
				Console.WriteLine("*** Signature Found ***")
				Console.WriteLine("Is valid: " & signature.IsValid)
				Console.WriteLine("Reason for signing: " & signature.Comments) ' This property is available in MS Word documents only.
				Console.WriteLine("Signature type: " & signature.SignatureType.ToString())
				Console.WriteLine("Time of signing: " & signature.SignTime)
				Console.WriteLine("Subject name: " & signature.CertificateHolder.Certificate.SubjectName.ToString())
				Console.WriteLine("Issuer name: " & signature.CertificateHolder.Certificate.IssuerName.Name)
				Console.WriteLine()
			Next signature
			'ExEnd

			Dim digitalSig As DigitalSignature = doc.DigitalSignatures(0)
			Assert.True(digitalSig.IsValid)
			Assert.AreEqual("Test Sign", digitalSig.Comments)
			Assert.AreEqual("XmlDsig", digitalSig.SignatureType.ToString())
			Assert.True(digitalSig.CertificateHolder.Certificate.Subject.Contains("Aspose Pty Ltd"))
			Assert.True(digitalSig.CertificateHolder.Certificate.IssuerName.Name IsNot Nothing AndAlso digitalSig.CertificateHolder.Certificate.IssuerName.Name.Contains("VeriSign"))
		End Sub

		' We don't include a sample certificate with the examples
		' so this exception is expected instead since the file is not there.
		<Test, ExpectedException(GetType(CryptographicException))> _
		Public Sub SignPdfDocument()
			'ExStart
			'ExFor:PdfSaveOptions
			'ExFor:PdfDigitalSignatureDetails
			'ExFor:PdfSaveOptions.DigitalSignatureDetails
			'ExFor:PdfDigitalSignatureDetails.#ctor(X509Certificate2, String, String, DateTime)
			'ExId:SignPDFDocument
			'ExSummary:Shows how to sign a generated PDF document using Aspose.Words.
			' Create a simple document from scratch.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)
			builder.Writeln("Test Signed PDF.")

			' Load the certificate from disk.
			' The other constructor overloads can be used to load certificates from different locations.
			Dim cert As New X509Certificate2(MyDir & "certificate.pfx", "feyb4lgcfbme")

			' Pass the certificate and details to the save options class to sign with.
			Dim options As New PdfSaveOptions()
			options.DigitalSignatureDetails = New PdfDigitalSignatureDetails(cert, "Test Signing", "Aspose Office", DateTime.Now)

			' Save the document as PDF with the digital signature set.
			doc.Save(MyDir & "Document.Signed Out.pdf", options)
			'ExEnd
		End Sub

		'This is for obfuscation bug WORDSNET-13036
		<Test, ExpectedException(GetType(TypeInitializationException))> _
		Public Sub SignDocument()
			Dim ch As CertificateHolder = CertificateHolder.Create(MyDir & "certificate.pfx", "123456")

			'By String
			Dim doc As New Document(MyDir & "TestRepeatingSection.doc")
			Dim outputDocFileName As String = MyDir & "\Artifacts\TestRepeatingSection.Signed.doc"

			DigitalSignatureUtil.Sign(doc.OriginalFileName, outputDocFileName, ch, "My comment", DateTime.Now)
		End Sub

		<Test> _
		Public Sub AppendAllDocumentsInFolder()
			Dim path As String = MyDir & "\Artifacts\Document.AppendDocumentsFromFolder.doc"

			' Delete the file that was created by the previous run as I don't want to append it again.
			If File.Exists(path) Then
				File.Delete(path)
			End If

			'ExStart
			'ExFor:Document.AppendDocument(Document, ImportFormatMode)
			'ExSummary:Shows how to use the AppendDocument method to combine all the documents in a folder to the end of a template document.
			' Lets start with a simple template and append all the documents in a folder to this document.
			Dim baseDoc As New Document()

			' Add some content to the template.
			Dim builder As New DocumentBuilder(baseDoc)
			builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1
			builder.Writeln("Template Document")
			builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal
			builder.Writeln("Some content here")

			' Gather the files which will be appended to our template document.
			' In this case we add the optional parameter to include the search only for files with the ".doc" extension.
			Dim files As New ArrayList(Directory.GetFiles(MyDir, "*.doc"))
			' The list of files may come in any order, let's sort the files by name so the documents are enumerated alphabetically.
			files.Sort()

			' Iterate through every file in the directory and append each one to the end of the template document.
			For Each fileName As String In files
				' We have some encrypted test documents in our directory, Aspose.Words can open encrypted documents 
				' but only with the correct password. Let's just skip them here for simplicity.
				Dim info As FileFormatInfo = FileFormatUtil.DetectFileFormat(fileName)
				If info.IsEncrypted Then
					Continue For
				End If

				Dim subDoc As New Document(fileName)
				baseDoc.AppendDocument(subDoc, ImportFormatMode.UseDestinationStyles)
			Next fileName

			' Save the combined document to disk.
			baseDoc.Save(path)
			'ExEnd
		End Sub

		<Test> _
		Public Sub JoinRunsWithSameFormatting()
			'ExStart
			'ExFor:Document.JoinRunsWithSameFormatting
			'ExSummary:Shows how to join runs in a document to reduce unneeded runs.
			' Let's load this particular document. It contains a lot of content that has been edited many times.
			' This means the document will most likely contain a large number of runs with duplicate formatting.
			Dim doc As New Document(MyDir & "Rendering.doc")

			' This is for illustration purposes only, remember how many run nodes we had in the original document.
			Dim runsBefore As Integer = doc.GetChildNodes(NodeType.Run, True).Count

			' Join runs with the same formatting. This is useful to speed up processing and may also reduce redundant
			' tags when exporting to HTML which will reduce the output file size.
			Dim joinCount As Integer = doc.JoinRunsWithSameFormatting()

			' This is for illustration purposes only, see how many runs are left after joining.
			Dim runsAfter As Integer = doc.GetChildNodes(NodeType.Run, True).Count

			Console.WriteLine("Number of runs before:{0}, after:{1}, joined:{2}", runsBefore, runsAfter, joinCount)

			' Save the optimized document to disk.
			doc.Save(MyDir & "\Artifacts\Document.JoinRunsWithSameFormatting.html")
			'ExEnd

			' Verify that runs were joined in the document.
			Assert.Less(runsAfter, runsBefore)
			Assert.AreNotEqual(0, joinCount)
		End Sub

		<Test> _
		Public Sub DetachTemplate()
			'ExStart
			'ExFor:Document.AttachedTemplate
			'ExSummary:Opens a document, makes sure it is no longer attached to a template and saves the document.
			Dim doc As New Document(MyDir & "Document.doc")
			doc.AttachedTemplate = ""
			doc.Save(MyDir & "\Artifacts\Document.DetachTemplate.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub DefaultTabStop()
			'ExStart
			'ExFor:Document.DefaultTabStop
			'ExFor:ControlChar.Tab
			'ExFor:ControlChar.TabChar
			'ExSummary:Changes default tab positions for the document and inserts text with some tab characters.
			Dim builder As New DocumentBuilder()

			' Set default tab stop to 72 points (1 inch).
			builder.Document.DefaultTabStop = 72

			builder.Writeln("Hello" & ControlChar.Tab & "World!")
			builder.Writeln("Hello" & ControlChar.TabChar & "World!")
			'ExEnd
		End Sub

		<Test> _
		Public Sub CloneDocument()
			'ExStart
			'ExFor:Document.Clone
			'ExId:CloneDocument
			'ExSummary:Shows how to deep clone a document.
			Dim doc As New Document(MyDir & "Document.doc")
			Dim clone As Document = doc.Clone()
			'ExEnd
		End Sub

		<Test> _
		Public Sub ChangeFieldUpdateCultureSource()
			' We will test this functionality creating a document with two fields with date formatting
			' field where the set language is different than the current culture, e.g German.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			' Insert content with German locale.
			builder.Font.LocaleId = 1031
			builder.InsertField("MERGEFIELD Date1 \@ ""dddd, d MMMM yyyy""")
			builder.Write(" - ")
			builder.InsertField("MERGEFIELD Date2 \@ ""dddd, d MMMM yyyy""")

			' Make sure that English culture is set then execute mail merge using current culture for
			' date formatting.
			Dim currentCulture As CultureInfo = Thread.CurrentThread.CurrentCulture
			Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
			doc.MailMerge.Execute(New String() { "Date1" }, New Object() { New DateTime(2011, 1, 01) })

			'ExStart
			'ExFor:Document.FieldOptions
			'ExFor:FieldOptions
			'ExFor:FieldOptions.FieldUpdateCultureSource
			'ExFor:FieldUpdateCultureSource
			'ExId:ChangeFieldUpdateCultureSource
			'ExSummary:Shows how to specify where the culture used for date formatting during field update and mail merge is chosen from.
			' Set the culture used during field update to the culture used by the field.
			doc.FieldOptions.FieldUpdateCultureSource = FieldUpdateCultureSource.FieldCode
			doc.MailMerge.Execute(New String() { "Date2" }, New Object() { New DateTime(2011, 1, 01) })
			'ExEnd

			' Verify the field update behaviour is correct.
			Assert.AreEqual("Saturday, 1 January 2011 - Samstag, 1 Januar 2011", doc.Range.Text.Trim())

			' Restore the original culture.
			Thread.CurrentThread.CurrentCulture = currentCulture
		End Sub

		<Test> _
		Public Sub ControlListLabelsExportToHtml()
			Dim doc As New Document(MyDir & "Lists.PrintOutAllLists.doc")
			Dim saveOptions As New HtmlSaveOptions(SaveFormat.Html)

			' This option uses <ul> and <ol> tags are used for list label representation if it doesn't cause formatting loss, 
			' otherwise HTML <p> tag is used. This is also the default value.
			saveOptions.ExportListLabels = ExportListLabels.Auto
			doc.Save(MyDir & "\Artifacts\Document.ExportListLabels Auto.html", saveOptions)

			' Using this option the <p> tag is used for any list label representation.
			saveOptions.ExportListLabels = ExportListLabels.AsInlineText
			doc.Save(MyDir & "\Artifacts\Document.ExportListLabels InlineText.html", saveOptions)

			' The <ul> and <ol> tags are used for list label representation. Some formatting loss is possible.
			saveOptions.ExportListLabels = ExportListLabels.ByHtmlTags
			doc.Save(MyDir & "\Artifacts\Document.ExportListLabels HtmlTags.html", saveOptions)
		End Sub

		<Test> _
		Public Sub DocumentGetText_ToString()
			'ExStart
			'ExFor:CompositeNode.GetText
			'ExFor:Node.ToString(SaveFormat)
			'ExId:NodeTxtExportDifferences
			'ExSummary:Shows the difference between calling the GetText and ToString methods on a node.
			Dim doc As New Document()

			' Enter a dummy field into the document.
			Dim builder As New DocumentBuilder(doc)
			builder.InsertField("MERGEFIELD Field")

			' GetText will retrieve all field codes and special characters
			Console.WriteLine("GetText() Result: " & doc.GetText())

			' ToString will export the node to the specified format. When converted to text it will not retrieve fields code 
			' or special characters, but will still contain some natural formatting characters such as paragraph markers etc. 
			' This is the same as "viewing" the document as if it was opened in a text editor.
			Console.WriteLine("ToString() Result: " & doc.ToString(SaveFormat.Text))
			'ExEnd
		End Sub

		<Test> _
		Public Sub DocumentByteArray()
			'ExStart
			'ExId:DocumentToFromByteArray
			'ExSummary:Shows how to convert a document object to an array of bytes and back into a document object again.
			' Load the document.
			Dim doc As New Document(MyDir & "Document.doc")

			' Create a new memory stream.
			Dim outStream As New MemoryStream()
			' Save the document to stream.
			doc.Save(outStream, SaveFormat.Docx)

			' Convert the document to byte form.
			Dim docBytes() As Byte = outStream.ToArray()

			' The bytes are now ready to be stored/transmitted.

			' Now reverse the steps to load the bytes back into a document object.
			Dim inStream As New MemoryStream(docBytes)

			' Load the stream into a new document object.
			Dim loadDoc As New Document(inStream)
			'ExEnd

			Assert.AreEqual(doc.GetText(), loadDoc.GetText())
		End Sub

		<Test> _
		Public Sub ProtectUnprotectDocument()
			'ExStart
			'ExFor:Document.Protect(ProtectionType,String)
			'ExId:ProtectDocument
			'ExSummary:Shows how to protect a document.
			Dim doc As New Document()
			doc.Protect(ProtectionType.AllowOnlyFormFields, "password")
			'ExEnd

			'ExStart
			'ExFor:Document.Unprotect
			'ExId:UnprotectDocument
			'ExSummary:Shows how to unprotect a document. Note that the password is not required.
			doc.Unprotect()
			'ExEnd

			'ExStart
			'ExFor:Document.Unprotect(String)
			'ExSummary:Shows how to unprotect a document using a password.
			doc.Unprotect("password")
			'ExEnd
		End Sub

		<Test> _
		Public Sub GetProtectionType()
			'ExStart
			'ExFor:Document.ProtectionType
			'ExId:GetProtectionType
			'ExSummary:Shows how to get protection type currently set in the document.
			Dim doc As New Document(MyDir & "Document.doc")
			Dim protectionType As ProtectionType = doc.ProtectionType
			'ExEnd
		End Sub

		<Test> _
		Public Sub DocumentEnsureMinimum()
			'ExStart
			'ExFor:Document.EnsureMinimum
			'ExSummary:Shows how to ensure the Document is valid (has the minimum nodes required to be valid).
			' Create a blank document then remove all nodes from it, the result will be a completely empty document.
			Dim doc As New Document()
			doc.RemoveAllChildren()

			' Ensure that the document is valid. Since the document has no nodes this method will create an empty section
			' and add an empty paragraph to make it valid.
			doc.EnsureMinimum()
			'ExEnd
		End Sub

		<Test> _
		Public Sub RemoveMacrosFromDocument()
			'ExStart
			'ExFor:Document.RemoveMacros
			'ExSummary:Shows how to remove all macros from a document.
			Dim doc As New Document(MyDir & "Document.doc")
			doc.RemoveMacros()
			'ExEnd
		End Sub

		<Test> _
		Public Sub UpdateTableLayout()
			'ExStart
			'ExFor:Document.UpdateTableLayout
			'ExId:UpdateTableLayout
			'ExSummary:Shows how to update the layout of tables in a document.
			Dim doc As New Document(MyDir & "Document.doc")

			' Normally this method is not necessary to call, as cell and table widths are maintained automatically.
			' This method may need to be called when exporting to PDF in rare cases when the table layout appears
			' incorrectly in the rendered output.
			doc.UpdateTableLayout()
			'ExEnd
		End Sub

		<Test> _
		Public Sub GetPageCount()
			'ExStart
			'ExFor:Document.PageCount
			'ExSummary:Shows how to invoke page layout and retrieve the number of pages in the document.
			Dim doc As New Document(MyDir & "Document.doc")

			' This invokes page layout which builds the document in memory so note that with large documents this
			' property can take time. After invoking this property, any rendering operation e.g rendering to PDF or image
			' will be instantaneous.
			Dim pageCount As Integer = doc.PageCount
			'ExEnd

			Assert.AreEqual(1, pageCount)
		End Sub

		<Test> _
		Public Sub UpdateFields()
			'ExStart
			'ExFor:Document.UpdateFields
			'ExId:UpdateFieldsInDocument
			'ExSummary:Shows how to update all fields in a document.
			Dim doc As New Document(MyDir & "Document.doc")
			doc.UpdateFields()
			'ExEnd
		End Sub

		<Test> _
		Public Sub GetUpdatedPageProperties()
			'ExStart
			'ExFor:Document.UpdateWordCount()
			'ExFor:BuiltInDocumentProperties.Characters
			'ExFor:BuiltInDocumentProperties.Words
			'ExFor:BuiltInDocumentProperties.Paragraphs
			'ExSummary:Shows how to update all list labels in a document.
			Dim doc As New Document(MyDir & "Document.doc")

			' Some work should be done here that changes the document's content.

			' Update the word, character and paragraph count of the document.
			doc.UpdateWordCount()

			' Display the updated document properties.
			Console.WriteLine("Characters: {0}", doc.BuiltInDocumentProperties.Characters)
			Console.WriteLine("Words: {0}", doc.BuiltInDocumentProperties.Words)
			Console.WriteLine("Paragraphs: {0}", doc.BuiltInDocumentProperties.Paragraphs)
			'ExEnd
		End Sub

		<Test, Explicit> _
		Public Sub TableStyleToDirectFormatting()
			'ExStart
			'ExFor:Document.ExpandTableStylesToDirectFormatting
			'ExId:TableStyleToDirectFormatting
			'ExSummary:Shows how to expand the formatting from styles onto the rows and cells of the table as direct formatting.
			Dim doc As New Document(MyDir & "Table.TableStyle.docx")

			' Get the first cell of the first table in the document.
			Dim table As Table = CType(doc.GetChild(NodeType.Table, 0, True), Table)
			Dim firstCell As Cell = table.FirstRow.FirstCell

			' First print the color of the cell shading. This should be empty as the current shading
			' is stored in the table style.
			Dim cellShadingBefore As Double = table.FirstRow.RowFormat.Height
			Console.WriteLine("Cell shading before style expansion: " & cellShadingBefore)

			' Expand table style formatting to direct formatting.
			doc.ExpandTableStylesToDirectFormatting()

			' Now print the cell shading after expanding table styles. A blue background pattern color
			' should have been applied from the table style.
			Dim cellShadingAfter As Double = table.FirstRow.RowFormat.Height
			Console.WriteLine("Cell shading after style expansion: " & cellShadingAfter)
			'ExEnd

			doc.Save(MyDir & "\Artifacts\Table.ExpandTableStyleFormatting.docx")

			Assert.AreEqual(Color.Empty, cellShadingBefore)
			Assert.AreNotEqual(Color.Empty, cellShadingAfter)
		End Sub

		<Test> _
		Public Sub GetOriginalFileInfo()
			'ExStart
			'ExFor:Document.OriginalFileName
			'ExFor:Document.OriginalLoadFormat
			'ExSummary:Shows how to retrieve the details of the path, filename and LoadFormat of a document from when the document was first loaded into memory.
			Dim doc As New Document(MyDir & "Document.doc")

			' This property will return the full path and file name where the document was loaded from.
			Dim originalFilePath As String = doc.OriginalFileName
			' Let's get just the file name from the full path.
			Dim originalFileName As String = Path.GetFileName(originalFilePath)

			' This is the original LoadFormat of the document.
			Dim loadFormat As LoadFormat = doc.OriginalLoadFormat
			'ExEnd
		End Sub

		<Test> _
		Public Sub RemoveSmartTagsFromDocument()
			'ExStart
			'ExFor:CompositeNode.RemoveSmartTags
			'ExSummary:Shows how to remove all smart tags from a document.
			Dim doc As New Document(MyDir & "Document.doc")
			doc.RemoveSmartTags()
			'ExEnd
		End Sub

		<Test> _
		Public Sub SetZoom()
			'ExStart
			'ExFor:Document.ViewOptions
			'ExFor:ViewOptions
			'ExFor:ViewOptions.ViewType
			'ExFor:ViewOptions.ZoomPercent
			'ExFor:ViewType
			'ExId:SetZoom
			'ExSummary:The following code shows how to make sure the document is displayed at 50% zoom when opened in Microsoft Word.
			Dim doc As New Document(MyDir & "Document.doc")
			doc.ViewOptions.ViewType = ViewType.PageLayout
			doc.ViewOptions.ZoomPercent = 50
			doc.Save(MyDir & "\Artifacts\Document.SetZoom.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub GetDocumentVariables()
			'ExStart
			'ExFor:Document.Variables
			'ExFor:VariableCollection
			'ExId:GetDocumentVariables
			'ExSummary:Shows how to enumerate over document variables.
			Dim doc As New Document(MyDir & "Document.doc")

			For Each entry As DictionaryEntry In doc.Variables
				Dim name As String = entry.Key.ToString()
				Dim value As String = entry.Value.ToString()

				' Do something useful.
				Console.WriteLine("Name: {0}, Value: {1}", name, value)
			Next entry
			'ExEnd
		End Sub

		<Test> _
		Public Sub FootnoteOptionsEx()
			'ExStart
			'ExFor:Document.FootnoteOptions
			'ExSummary:Shows how to insert a footnote and apply footnote options.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			builder.InsertFootnote(FootnoteType.Footnote, "My Footnote.")

			' Change your document's footnote options.
			doc.FootnoteOptions.Location = FootnoteLocation.BottomOfPage
			doc.FootnoteOptions.NumberStyle = NumberStyle.Arabic
			doc.FootnoteOptions.StartNumber = 1

			doc.Save(MyDir & "\Artifacts\Document.FootnoteOptions.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub CompareEx()
			'ExStart
			'ExFor:Document.Compare
			'ExSummary:Shows how to apply the compare method to two documents and then use the results. 
			Dim doc1 As New Document(MyDir & "Document.Compare.1.doc")
			Dim doc2 As New Document(MyDir & "Document.Compare.2.doc")

			' If either document has a revision, an exception will be thrown.
			If doc1.Revisions.Count = 0 AndAlso doc2.Revisions.Count = 0 Then
				doc1.Compare(doc2, "authorName", DateTime.Now)
			End If

			' If doc1 and doc2 are different, doc1 now has some revisons after the comparison, which can now be viewed and processed.
			For Each r As Revision In doc1.Revisions
				Console.WriteLine(r.RevisionType)
			Next r

			' All the revisions in doc1 are differences between doc1 and doc2, so accepting them on doc1 transforms doc1 into doc2.
			doc1.Revisions.AcceptAll()

			' doc1, when saved, now resembles doc2.
			doc1.Save(MyDir & "\Artifacts\Document.CompareEx.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub RemoveExternalSchemaReferencesEx()
			'ExStart
			'ExFor:Document.RemoveExternalSchemaReferences
			'ExSummary:Shows how to remove all external XML schema references from a document. 
			Dim doc As New Document(MyDir & "Document.doc")
			doc.RemoveExternalSchemaReferences()
			'ExEnd
		End Sub

		<Test> _
		Public Sub RemoveUnusedResourcesEx()
			'ExStart
			'ExFor:Document.RemoveUnusedResources
			'ExSummary:Shows how to remove all unused styles and lists from a document. 
			Dim doc As New Document(MyDir & "Document.doc")
			doc.RemoveUnusedResources()
			'ExEnd
		End Sub

		<Test> _
		Public Sub StartTrackRevisionsEx()
			'ExStart
			'ExFor:Document.StartTrackRevisions(String)
			'ExFor:Document.StartTrackRevisions(String, DateTime)
			'ExFor:Document.StopTrackRevisions
			'ExSummary:Shows how tracking revisions affects document editing. 
			Dim doc As New Document()

			' This text will appear as normal text in the document and no revisions will be counted.
			doc.FirstSection.Body.FirstParagraph.Runs.Add(New Run(doc, "Hello world!"))
			Console.WriteLine(doc.Revisions.Count) ' 0

			doc.StartTrackRevisions("Author")

			' This text will appear as a revision. 
			' We did not specify a time while calling StartTrackRevisions(), so the date/time that's noted
			' on the revision will be the real time when StartTrackRevisions() executes.
			doc.FirstSection.Body.AppendParagraph("Hello again!")
			Console.WriteLine(doc.Revisions.Count) ' 2

			' Stopping the tracking of revisions makes this text appear as normal text. 
			' Revisions are not counted when the document is changed.
			doc.StopTrackRevisions()
			doc.FirstSection.Body.AppendParagraph("Hello again!")
			Console.WriteLine(doc.Revisions.Count) ' 2

			' Specifying some date/time will apply that date/time to all subsequent revisions until StopTrackRevisions() is called.
			' Note that placing values such as DateTime.MinValue as an argument will create revisions that do not have a date/time at all.
			doc.StartTrackRevisions("Author", New DateTime(1970, 1, 1))
			doc.FirstSection.Body.AppendParagraph("Hello again!")
			Console.WriteLine(doc.Revisions.Count) ' 4

			doc.Save(MyDir & "\Artifacts\Document.StartTrackRevisions.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub AcceptAllRevisions()
			'ExStart
			'ExFor:Document.AcceptAllRevisions
			'ExSummary:Shows how to accept all tracking changes in the document.
			Dim doc As New Document(MyDir & "Document.doc")

			' Start tracking and make some revisions.
			doc.StartTrackRevisions("Author")
			doc.FirstSection.Body.AppendParagraph("Hello world!")

			' Revisions will now show up as normal text in the output document.
			doc.AcceptAllRevisions()
			doc.Save(MyDir & "\Artifacts\Document.AcceptedRevisions.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub UpdateThumbnailEx()
			'ExStart
			'ExFor:Document.UpdateThumbnail()
			'ExFor:Document.UpdateThumbnail(ThumbnailGeneratingOptions)
			'ExSummary:Shows how to update a document's thumbnail.
			Dim doc As New Document()

			' Update document's thumbnail the default way. 
			doc.UpdateThumbnail()

			' Review/change thumbnail options and then update document's thumbnail.
			Dim tgo As New ThumbnailGeneratingOptions()

			Console.WriteLine("Thumbnail size: {0}", tgo.ThumbnailSize)
			tgo.GenerateFromFirstPage = True

			doc.UpdateThumbnail(tgo)
			'ExEnd
		End Sub

		'For assert this test you need to open "HyphenationOptions OUT.docx" and check that hyphen are added in the end of the first line
		<Test> _
		Public Sub HyphenationOptions()
			Dim doc As New Document()

			DocumentHelper.InsertNewRun(doc, "poqwjopiqewhpefobiewfbiowefob ewpj weiweohiewobew ipo efoiewfihpewfpojpief pijewfoihewfihoewfphiewfpioihewfoihweoihewfpj", 0)

			doc.HyphenationOptions.AutoHyphenation = True
			doc.HyphenationOptions.ConsecutiveHyphenLimit = 2
			doc.HyphenationOptions.HyphenationZone = 720 ' 0.5 inch
			doc.HyphenationOptions.HyphenateCaps = True

			Dim dstStream As New MemoryStream()
			doc.Save(dstStream, SaveFormat.Docx)

			Assert.AreEqual(True, doc.HyphenationOptions.AutoHyphenation)
			Assert.AreEqual(2, doc.HyphenationOptions.ConsecutiveHyphenLimit)
			Assert.AreEqual(720, doc.HyphenationOptions.HyphenationZone)
			Assert.AreEqual(True, doc.HyphenationOptions.HyphenateCaps)

			doc.Save(MyDir & "HyphenationOptions.docx")
		End Sub

		<Test> _
		Public Sub HyphenationOptionsDefaultValues()
			Dim doc As New Document()

			Dim dstStream As New MemoryStream()
			doc.Save(dstStream, SaveFormat.Docx)

			Assert.AreEqual(False, doc.HyphenationOptions.AutoHyphenation)
			Assert.AreEqual(0, doc.HyphenationOptions.ConsecutiveHyphenLimit)
			Assert.AreEqual(360, doc.HyphenationOptions.HyphenationZone) ' 0.25 inch
			Assert.AreEqual(True, doc.HyphenationOptions.HyphenateCaps)
		End Sub

		<Test, TestCase(0, 0, ExpectedException := GetType(ArgumentOutOfRangeException)), TestCase(-1, 360, ExpectedException := GetType(ArgumentOutOfRangeException))> _
		Public Sub HyphenationOptionsExceptions(ByVal consecutiveHyphenLimit As Integer, ByVal hyphenationZone As Integer)
			Dim doc As New Document()

			doc.HyphenationOptions.ConsecutiveHyphenLimit = consecutiveHyphenLimit
			doc.HyphenationOptions.HyphenationZone = hyphenationZone
		End Sub

		<Test> _
		Public Sub ExtractPlainTextFromDocument()
			'ExStart
			'ExFor:Document.ExtractText(string)
			'ExFor:Document.ExtractText(string, LoadOptions)
			'ExFor:PlaintextDocument.Text
			'ExFor:PlaintextDocument.BuiltInDocumentProperties
			'ExFor:PlaintextDocument.CustomDocumentProperties
			'ExSummary:Shows how to extract plain text from the document and get it properties
			Dim plaintext As New PlainTextDocument(MyDir & "Bookmark.doc")
			Assert.AreEqual("This is a bookmarked text." & Constants.vbFormFeed, plaintext.Text)

			Dim loadOptions As New LoadOptions()
			loadOptions.AllowTrailingWhitespaceForListItems = False

			plaintext = New PlainTextDocument(MyDir & "Bookmark.doc", loadOptions)
			Assert.AreEqual("This is a bookmarked text." & Constants.vbFormFeed, plaintext.Text)

			Dim builtInDocumentProperties As BuiltInDocumentProperties = plaintext.BuiltInDocumentProperties
			Assert.AreEqual("Aspose", builtInDocumentProperties.Company)

			Dim customDocumentProperties As CustomDocumentProperties = plaintext.CustomDocumentProperties
			Assert.IsEmpty(customDocumentProperties)
			'ExEnd
		End Sub

		<Test> _
		Public Sub ExtractPlainTextFromStream()
			'ExStart
			'ExFor:Document.ExtractText(Stream)
			'ExFor:Document.ExtractText(Stream, LoadOptions)
			'ExSummary:
			Dim docStream As Stream = New FileStream(MyDir & "Bookmark.doc", FileMode.Open)

			Dim plaintext As New PlainTextDocument(docStream)
			Assert.AreEqual("This is a bookmarked text." & Constants.vbFormFeed, plaintext.Text)

			docStream.Close()

			docStream = New FileStream(MyDir & "Bookmark.doc", FileMode.Open)

			Dim loadOptions As New LoadOptions()
			loadOptions.AllowTrailingWhitespaceForListItems = False

			plaintext = New PlainTextDocument(docStream, loadOptions)
			Assert.AreEqual("This is a bookmarked text." & Constants.vbFormFeed, plaintext.Text)

			docStream.Close()
			'ExEnd
		End Sub

		<Test> _
		Public Sub GetShapeAltTextTitle()
			'ExStart
			'ExFor:Shape.Title
			'ExSummary:Shows how to get or set alt text title for shape object
			Dim doc As New Document()

			' Create textbox shape.
			Dim shape As New Shape(doc, ShapeType.Cube)
			shape.Width = 431.5
			shape.Height = 346.35
			shape.Title = "Alt Text Title"

			Dim paragraph As New Paragraph(doc)
			paragraph.AppendChild(New Run(doc, "Test"))

			' Insert paragraph into the textbox.
			shape.AppendChild(paragraph)

			' Insert textbox into the document.
			doc.FirstSection.Body.FirstParagraph.AppendChild(shape)

			Dim dstStream As New MemoryStream()
			doc.Save(dstStream, SaveFormat.Docx)

			Dim shapes() As Node = doc.GetChildNodes(NodeType.Shape, True).ToArray()
			shape = CType(shapes(0), Shape)

			Assert.AreEqual("Alt Text Title", shape.Title)
			'ExEnd
		End Sub

		<Test> _
		Public Sub GetOrSetDocumentThemeProperties()
			Dim doc As New Document()

			Dim theme As Theme = doc.Theme

			theme.Colors.Accent1 = Color.Black
			theme.Colors.Dark1 = Color.Blue
			theme.Colors.FollowedHyperlink = Color.White
			theme.Colors.Hyperlink = Color.WhiteSmoke
			theme.Colors.Light1 = Color.Empty 'There is default Color.Black

			theme.MajorFonts.ComplexScript = "Arial"
			theme.MajorFonts.EastAsian = String.Empty
			theme.MajorFonts.Latin = "Times New Roman"

			theme.MinorFonts.ComplexScript = String.Empty
			theme.MinorFonts.EastAsian = "Times New Roman"
			theme.MinorFonts.Latin = "Arial"

			Dim dstStream As New MemoryStream()
			doc.Save(dstStream, SaveFormat.Docx)

			Assert.AreEqual(Color.Black.ToArgb(), doc.Theme.Colors.Accent1.ToArgb())
			Assert.AreEqual(Color.Blue.ToArgb(), doc.Theme.Colors.Dark1.ToArgb())
			Assert.AreEqual(Color.White.ToArgb(), doc.Theme.Colors.FollowedHyperlink.ToArgb())
			Assert.AreEqual(Color.WhiteSmoke.ToArgb(), doc.Theme.Colors.Hyperlink.ToArgb())
			Assert.AreEqual(Color.Black.ToArgb(), doc.Theme.Colors.Light1.ToArgb())

			Assert.AreEqual("Arial", doc.Theme.MajorFonts.ComplexScript)
			Assert.AreEqual(String.Empty, doc.Theme.MajorFonts.EastAsian)
			Assert.AreEqual("Times New Roman", doc.Theme.MajorFonts.Latin)

			Assert.AreEqual(String.Empty, doc.Theme.MinorFonts.ComplexScript)
			Assert.AreEqual("Times New Roman", doc.Theme.MinorFonts.EastAsian)
			Assert.AreEqual("Arial", doc.Theme.MinorFonts.Latin)
		End Sub
	End Class
End Namespace