' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports System.Globalization
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Threading
Imports Aspose.BarCodeRecognition
Imports Aspose.Words
Imports Aspose.Words.Fields

Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Public Class ExField
		Inherits ApiExampleBase
		<Test> _
		Public Sub UpdateToc()
			Dim doc As New Document()

			'ExStart
			'ExId:UpdateTOC
			'ExSummary:Shows how to completely rebuild TOC fields in the document by invoking field update.
			doc.UpdateFields()
			'ExEnd
		End Sub

		<Test> _
		Public Sub GetFieldType()
			Dim doc As New Document(MyDir & "Document.TableOfContents.doc")

			'ExStart
			'ExFor:FieldType
			'ExFor:FieldChar
			'ExFor:FieldChar.FieldType
			'ExSummary:Shows how to find the type of field that is represented by a node which is derived from FieldChar.
			Dim fieldStart As FieldChar = CType(doc.GetChild(NodeType.FieldStart, 0, True), FieldChar)
			Dim type As FieldType = fieldStart.FieldType
			'ExEnd
		End Sub

		<Test> _
		Public Sub GetFieldFromDocument()
			'ExStart
			'ExFor:FieldChar.GetField
			'ExId:GetField
			'ExSummary:Demonstrates how to retrieve the field class from an existing FieldStart node in the document.
			Dim doc As New Document(MyDir & "Document.TableOfContents.doc")

			Dim fieldStart As FieldStart = CType(doc.GetChild(NodeType.FieldStart, 0, True), FieldStart)

			' Retrieve the facade object which represents the field in the document.
			Dim field As Field = fieldStart.GetField()

			Console.WriteLine("Field code:" & field.GetFieldCode())
			Console.WriteLine("Field result: " & field.Result)
			Console.WriteLine("Is locked: " & field.IsLocked)

			' This updates only this field in the document.
			field.Update()
			'ExEnd
		End Sub

		<Test> _
		Public Sub GetFieldFromFieldCollection()
			'ExStart
			'ExId:GetFieldFromFieldCollection
			'ExSummary:Demonstrates how to retrieve a field using the range of a node.
			Dim doc As New Document(MyDir & "Document.TableOfContents.doc")

			Dim field As Field = doc.Range.Fields(0)

			' This should be the first field in the document - a TOC field.
			Console.WriteLine(field.Type)
			'ExEnd
		End Sub

		<Test> _
		Public Sub InsertTcField()
			'ExStart
			'ExId:InsertTCField
			'ExSummary:Shows how to insert a TC field into the document using DocumentBuilder.
			' Create a blank document.
			Dim doc As New Document()

			' Create a document builder to insert content with.
			Dim builder As New DocumentBuilder(doc)

			' Insert a TC field at the current document builder position.
			builder.InsertField("TC ""Entry Text"" \f t")
			'ExEnd
		End Sub

		<Test> _
		Public Sub ChangeLocale()
			' Create a blank document.
			Dim doc As New Document()
			Dim b As New DocumentBuilder(doc)
			b.InsertField("MERGEFIELD Date")

			'ExStart
			'ExId:ChangeCurrentCulture
			'ExSummary:Shows how to change the culture used in formatting fields during update.
			' Store the current culture so it can be set back once mail merge is complete.
			Dim currentCulture As CultureInfo = Thread.CurrentThread.CurrentCulture
			' Set to German language so dates and numbers are formatted using this culture during mail merge.
			Thread.CurrentThread.CurrentCulture = New CultureInfo("de-DE")

			' Execute mail merge.
			doc.MailMerge.Execute(New String() {"Date"}, New Object() {DateTime.Now})

			' Restore the original culture.
			Thread.CurrentThread.CurrentCulture = currentCulture
			'ExEnd

			doc.Save(MyDir & "\Artifacts\Field.ChangeLocale.doc")
		End Sub

		<Test> _
		Public Sub RemoveTocFromDocument()
			'ExStart
			'ExFor:CompositeNode.GetChildNodes(NodeType, Boolean)
			'ExId:RemoveTableOfContents
			'ExSummary:Demonstrates how to remove a specified TOC from a document.
			' Open a document which contains a TOC.
			Dim doc As New Document(MyDir & "Document.TableOfContents.doc")

			' Remove the first TOC from the document.
			Dim tocField As Field = doc.Range.Fields(0)
			tocField.Remove()

			' Save the output.
			doc.Save(MyDir & "\Artifacts\Document.TableOfContentsRemoveTOC.doc")
			'ExEnd
		End Sub

		'ExStart
		'ExId:TCFieldsRangeReplace
		'ExSummary:Shows how to find and insert a TC field at text in a document. 
		<Test> _
		Public Sub InsertTcFieldsAtText()
			Dim doc As New Document()

			' Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document.
			doc.Range.Replace(New Regex("The Beginning"), New InsertTcFieldHandler("Chapter 1", "\l 1"), False)
		End Sub

		Public Class InsertTcFieldHandler
			Implements IReplacingCallback
			' Store the text and switches to be used for the TC fields.
			Private mFieldText As String
			Private mFieldSwitches As String

			''' <summary>
			''' The switches to use for each TC field. Can be an empty string or null.
			''' </summary>
			Public Sub New(ByVal switches As String)
				Me.New(String.Empty, switches)
				Me.mFieldSwitches = switches
			End Sub

			''' <summary>
			''' The display text and switches to use for each TC field. Display name can be an empty string or null.
			''' </summary>
			Public Sub New(ByVal text As String, ByVal switches As String)
				Me.mFieldText = text
				Me.mFieldSwitches = switches
			End Sub

			Private Function IReplacingCallback_Replacing(ByVal args As ReplacingArgs) As ReplaceAction Implements IReplacingCallback.Replacing
				' Create a builder to insert the field.
				Dim builder As New DocumentBuilder(CType(args.MatchNode.Document, Document))
				' Move to the first node of the match.
				builder.MoveTo(args.MatchNode)

				' If the user specified text to be used in the field as display text then use that, otherwise use the 
				' match string as the display text.
				Dim insertText As String

				If (Not String.IsNullOrEmpty(Me.mFieldText)) Then
					insertText = Me.mFieldText
				Else
					insertText = args.Match.Value
				End If

				' Insert the TC field before this node using the specified string as the display text and user defined switches.
				builder.InsertField(String.Format("TC ""{0}"" {1}", insertText, Me.mFieldSwitches))

				' We have done what we want so skip replacement.
				Return ReplaceAction.Skip
			End Function
		End Class

		'ExEnd

		'Bug: there is no isAfter parameter at BuildAndInsert (exception), need more info from dev
		<Test> _
		Public Sub InsertFieldWithFieldBuilder()
			Dim doc As New Document()

			'Add some text into the paragraph
			Dim run As Run = DocumentHelper.InsertNewRun(doc, " Hello World!", 0)

			Dim argumentBuilder As New FieldArgumentBuilder()
			argumentBuilder.AddField(New FieldBuilder(FieldType.FieldMergeField))
			argumentBuilder.AddText("BestField")

			Dim fieldBuilder As New FieldBuilder(FieldType.FieldIf)
			fieldBuilder.AddArgument(argumentBuilder).AddArgument("=").AddArgument("BestField").AddArgument(10).AddArgument(20.0).AddSwitch("12", "13").BuildAndInsert(run)

			doc.UpdateFields()
		End Sub

		<Test, ExpectedException(GetType(ArgumentException), ExpectedMessage := "Cannot add a node before/after itself.")> _
		Public Sub InsertFieldWithFieldBuilderException()
			Dim doc As New Document()

			'Add some text into the paragraph
			Dim run As Run = DocumentHelper.InsertNewRun(doc, " Hello World!", 0)

			Dim argumentBuilder As New FieldArgumentBuilder()
			argumentBuilder.AddField(New FieldBuilder(FieldType.FieldMergeField))
			argumentBuilder.AddNode(run)
			argumentBuilder.AddText("Text argument builder")

			Dim fieldBuilder As New FieldBuilder(FieldType.FieldIncludeText)
			fieldBuilder.AddArgument(argumentBuilder).AddArgument("=").AddArgument("BestField").AddArgument(10).AddArgument(20.0).BuildAndInsert(run)
		End Sub

		<Test> _
		Public Sub BarCodeWord2Pdf()
			Dim doc As New Document(MyDir & "BarCode.docx")

			' Set custom barcode generator
			doc.FieldOptions.BarcodeGenerator = New CustomBarcodeGenerator()

			doc.Save(MyDir & "\Artifacts\BarCode.pdf")

			Dim barCode As BarCodeReader = BarCodeReaderPdf(MyDir & "\Artifacts\BarCode.pdf")
			Assert.AreEqual("QR", barCode.GetReadType().ToString())
		End Sub

		Private Function BarCodeReaderPdf(ByVal filename As String) As BarCodeReader
			'Set license for Aspose.BarCode
			Dim licenceBarCode As New Aspose.BarCode.License()

			licenceBarCode.SetLicense("X:\awnet\TestData\Licenses\Aspose.Total.lic")

			'bind the pdf document
			Dim pdfExtractor As New Aspose.Pdf.Facades.PdfExtractor()
			pdfExtractor.BindPdf(filename)

			'set page range for image extraction
			pdfExtractor.StartPage = 1
			pdfExtractor.EndPage = 1

			pdfExtractor.ExtractImage()

			'save image to stream
			Dim imageStream As New MemoryStream()
			pdfExtractor.GetNextImage(imageStream)
			imageStream.Position = 0

			'recognize the barcode from the image stream above
			Dim barcodeReader As New BarCodeReader(imageStream, BarCodeReadType.QR)
			Do While barcodeReader.Read()
				Console.WriteLine("Codetext found: " & barcodeReader.GetCodeText() & ", Symbology: " & barcodeReader.GetReadType())
			Loop

			'close the reader
			barcodeReader.Close()

			Return barcodeReader
		End Function

		'For assert result of the test you need to open "UpdateFieldIgnoringMergeFormat Out.docx" and check that image are added correct and without truncated inside frame
		<Test> _
		Public Sub UpdateFieldIgnoringMergeFormat()
			'ExStart
			'ExFor:FieldIncludePicture.Update(Bool)
			'ExSummary:Shows a way to update a field ignoring the MERGEFORMAT switch
			Dim loadOptions As New LoadOptions()
			loadOptions.PreserveIncludePictureField = True

			Dim doc As New Document(MyDir & "UpdateFieldIgnoringMergeFormat.docx", loadOptions)

			For Each field As Field In doc.Range.Fields
				If field.Type.Equals(FieldType.FieldIncludePicture) Then
					Dim includePicture As FieldIncludePicture = CType(field, FieldIncludePicture)

					includePicture.SourceFullName = MyDir & "\Images\dotnet-logo.png"
					includePicture.Update(True)
				End If
			Next field

			doc.UpdateFields()
			doc.Save(MyDir & "UpdateFieldIgnoringMergeFormat Out.docx")
			'ExEnd
		End Sub
	End Class
End Namespace
