' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Threading

Imports Aspose.Words

Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Public Class ExSection
		Inherits ApiExampleBase
		<Test> _
		Public Sub Protect()
			'ExStart
			'ExFor:Document.Protect(ProtectionType)
			'ExFor:ProtectionType
			'ExFor:Section.ProtectedForForms
			'ExSummary:Protects a section so only editing in form fields is possible.
			' Create a blank document
			Dim doc As New Document()

			' Insert two sections with some text
			Dim builder As New DocumentBuilder(doc)
			builder.Writeln("Section 1. Unprotected.")
			builder.InsertBreak(BreakType.SectionBreakContinuous)
			builder.Writeln("Section 2. Protected.")

			' Section protection only works when document protection is turned and only editing in form fields is allowed.
			doc.Protect(ProtectionType.AllowOnlyFormFields)

			' By default, all sections are protected, but we can selectively turn protection off.
			doc.Sections(0).ProtectedForForms = False

			builder.Document.Save(MyDir & "\Artifacts\Section.Protect.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub AddRemove()
			'ExStart
			'ExFor:Document.Sections
			'ExFor:Section.Clone
			'ExFor:SectionCollection
			'ExFor:NodeCollection.RemoveAt(Int32)
			'ExSummary:Shows how to add/remove sections in a document.
			' Open the document.
			Dim doc As New Document(MyDir & "Section.AddRemove.doc")

			' This shows what is in the document originally. The document has two sections.
			Console.WriteLine(doc.GetText())

			' Delete the first section from the document
			doc.Sections.RemoveAt(0)

			' Duplicate the last section and append the copy to the end of the document.
			Dim lastSectionIdx As Integer = doc.Sections.Count - 1
			Dim newSection As Section = doc.Sections(lastSectionIdx).Clone()
			doc.Sections.Add(newSection)

			' Check what the document contains after we changed it.
			Console.WriteLine(doc.GetText())
			'ExEnd

			Assert.AreEqual("Hello2" & Constants.vbFormFeed & "Hello2" & Constants.vbFormFeed, doc.GetText())
		End Sub

		<Test> _
		Public Sub CreateFromScratch()
			'ExStart
			'ExFor:Node.GetText
			'ExFor:CompositeNode.RemoveAllChildren
			'ExFor:CompositeNode.AppendChild
			'ExFor:Section
			'ExFor:Section.#ctor
			'ExFor:Section.PageSetup
			'ExFor:PageSetup.SectionStart
			'ExFor:PageSetup.PaperSize
			'ExFor:SectionStart
			'ExFor:PaperSize
			'ExFor:Body
			'ExFor:Body.#ctor
			'ExFor:Paragraph
			'ExFor:Paragraph.#ctor
			'ExFor:Paragraph.ParagraphFormat
			'ExFor:ParagraphFormat
			'ExFor:ParagraphFormat.StyleName
			'ExFor:ParagraphFormat.Alignment
			'ExFor:ParagraphAlignment
			'ExFor:Run
			'ExFor:Run.#ctor(DocumentBase)
			'ExFor:Run.Text
			'ExFor:Inline.Font
			'ExSummary:Creates a simple document from scratch using the Aspose.Words object model.

			' Create an "empty" document. Note that like in Microsoft Word, 
			' the empty document has one section, body and one paragraph in it.
			Dim doc As New Document()

			' This truly makes the document empty. No sections (not possible in Microsoft Word).
			doc.RemoveAllChildren()

			' Create a new section node. 
			' Note that the section has not yet been added to the document, 
			' but we have to specify the parent document.
			Dim section As New Section(doc)

			' Append the section to the document.
			doc.AppendChild(section)

			' Lets set some properties for the section.
			section.PageSetup.SectionStart = SectionStart.NewPage
			section.PageSetup.PaperSize = PaperSize.Letter


			' The section that we created is empty, lets populate it. The section needs at least the Body node.
			Dim body As New Body(doc)
			section.AppendChild(body)


			' The body needs to have at least one paragraph.
			' Note that the paragraph has not yet been added to the document, 
			' but we have to specify the parent document.
			' The parent document is needed so the paragraph can correctly work
			' with styles and other document-wide information.
			Dim para As New Paragraph(doc)
			body.AppendChild(para)

			' We can set some formatting for the paragraph
			para.ParagraphFormat.StyleName = "Heading 1"
			para.ParagraphFormat.Alignment = ParagraphAlignment.Center


			' So far we have one empty paragraph in the document.
			' The document is valid and can be saved, but lets add some text before saving.
			' Create a new run of text and add it to our paragraph.
			Dim run As New Run(doc)
			run.Text = "Hello World!"
			run.Font.Color = Color.Red
			para.AppendChild(run)


			' As a matter of interest, you can retrieve text of the whole document and
			' see that \x000c is automatically appended. \x000c is the end of section character.
			Console.WriteLine("Hello World!" & Constants.vbFormFeed, doc.GetText())

			' Save the document.
			doc.Save(MyDir & "\Artifacts\Section.CreateFromScratch.doc")
			'ExEnd

			Assert.AreEqual("Hello World!" & Constants.vbFormFeed, doc.GetText())
		End Sub

		<Test> _
		Public Sub EnsureSectionMinimum()
			'ExStart
			'ExFor:Section.EnsureMinimum
			'ExSummary:Ensures that a section is valid.
			' Create a blank document
			Dim doc As New Document()
			Dim section As Section = doc.FirstSection

			' Makes sure that the section contains a body with at least one paragraph.
			section.EnsureMinimum()
			'ExEnd
		End Sub

		<Test> _
		Public Sub BodyEnsureMinimum()
			'ExStart
			'ExFor:Section.Body
			'ExFor:Body.EnsureMinimum
			'ExSummary:Clears main text from all sections from the document leaving the sections themselves.

			' Open a document.
			Dim doc As New Document(MyDir & "Section.BodyEnsureMinimum.doc")

			' This shows what is in the document originally. The document has two sections.
			Console.WriteLine(doc.GetText())

			' Loop through all sections in the document.
			For Each section As Section In doc.Sections
				' Each section has a Body node that contains main story (main text) of the section.
				Dim body As Body = section.Body

				' This clears all nodes from the body.
				body.RemoveAllChildren()

				' Technically speaking, for the main story of a section to be valid, it needs to have
				' at least one empty paragraph. That's what the EnsureMinimum method does.
				body.EnsureMinimum()
			Next section

			' Check how the content of the document looks now.
			Console.WriteLine(doc.GetText())
			'ExEnd

			Assert.AreEqual(Constants.vbFormFeed + Constants.vbFormFeed, doc.GetText())
		End Sub

		<Test> _
		Public Sub BodyNodeType()
			'ExStart
			'ExFor:Body.NodeType
			'ExFor:HeaderFooter.NodeType
			'ExFor:Document.FirstSection
			'ExSummary:Shows how you can enumerate through children of a composite node and detect types of the children nodes.

			' Open a document.
			Dim doc As New Document(MyDir & "Section.BodyNodeType.doc")

			' Get the first section in the document.
			Dim section As Section = doc.FirstSection

			' A Section is a composite node and therefore can contain child nodes.
			' Section can contain only Body and HeaderFooter nodes.
			For Each node As Node In section
				' Every node has the NodeType property.
				Select Case node.NodeType
					Case NodeType.Body
						' If the node type is Body, we can cast the node to the Body class.
						Dim body As Body = CType(node, Body)

						' Write the content of the main story of the section to the console.
						Console.WriteLine("*** Body ***")
						Console.WriteLine(body.GetText())
						Exit Select
					Case NodeType.HeaderFooter
						' If the node type is HeaderFooter, we can cast the node to the HeaderFooter class.
						Dim headerFooter As HeaderFooter = CType(node, HeaderFooter)

						' Write the content of the header footer to the console.
						Console.WriteLine("*** HeaderFooter ***")
						Console.WriteLine(headerFooter.HeaderFooterType)
						Console.WriteLine(headerFooter.GetText())
						Exit Select
					Case Else
						' Other types of nodes never occur inside a Section node.
						Throw New Exception("Unexpected node type in a section.")
				End Select
			Next node
			'ExEnd
		End Sub

		<Test> _
		Public Sub SectionsAccessByIndex()
			'ExStart
			'ExFor:SectionCollection.Item(Int32)
			'ExId:SectionsAccessByIndex
			'ExSummary:Shows how to access a section at the specified index.
			Dim doc As New Document(MyDir & "Document.doc")
			Dim section As Section = doc.Sections(0)
			'ExEnd
		End Sub

		<Test> _
		Public Sub SectionsAddSection()
			'ExStart
			'ExFor:NodeCollection.Add
			'ExId:SectionsAddSection
			'ExSummary:Shows how to add a section to the end of the document.
			Dim doc As New Document(MyDir & "Document.doc")
			Dim sectionToAdd As New Section(doc)
			doc.Sections.Add(sectionToAdd)
			'ExEnd
		End Sub

		<Test> _
		Public Sub SectionsDeleteSection()
			'ExStart
			'ExId:SectionsDeleteSection
			'ExSummary:Shows how to remove a section at the specified index.
			Dim doc As New Document(MyDir & "Document.doc")
			doc.Sections.RemoveAt(0)
			'ExEnd
		End Sub

		<Test> _
		Public Sub SectionsDeleteAllSections()
			'ExStart
			'ExFor:NodeCollection.Clear
			'ExId:SectionsDeleteAllSections
			'ExSummary:Shows how to remove all sections from a document.
			Dim doc As New Document(MyDir & "Document.doc")
			doc.Sections.Clear()
			'ExEnd
		End Sub

		<Test> _
		Public Sub SectionsAppendSectionContent()
			'ExStart
			'ExFor:Section.AppendContent
			'ExFor:Section.PrependContent
			'ExId:SectionsAppendSectionContent
			'ExSummary:Shows how to append content of an existing section. The number of sections in the document remains the same.
			Dim doc As New Document(MyDir & "Section.AppendContent.doc")

			' This is the section that we will append and prepend to.
			Dim section As Section = doc.Sections(2)

			' This copies content of the 1st section and inserts it at the beginning of the specified section.
			Dim sectionToPrepend As Section = doc.Sections(0)
			section.PrependContent(sectionToPrepend)

			' This copies content of the 2nd section and inserts it at the end of the specified section.
			Dim sectionToAppend As Section = doc.Sections(1)
			section.AppendContent(sectionToAppend)
			'ExEnd
		End Sub

		<Test> _
		Public Sub SectionsDeleteSectionContent()
			'ExStart
			'ExFor:Section.ClearContent
			'ExId:SectionsDeleteSectionContent
			'ExSummary:Shows how to delete main content of a section.
			Dim doc As New Document(MyDir & "Document.doc")
			Dim section As Section = doc.Sections(0)
			section.ClearContent()
			'ExEnd
		End Sub

		<Test> _
		Public Sub SectionsDeleteHeaderFooter()
			'ExStart
			'ExFor:Section.ClearHeadersFooters
			'ExId:SectionsDeleteHeaderFooter
			'ExSummary:Clears content of all headers and footers in a section.
			Dim doc As New Document(MyDir & "Document.doc")
			Dim section As Section = doc.Sections(0)
			section.ClearHeadersFooters()
			'ExEnd
		End Sub

		<Test> _
		Public Sub SectionDeleteHeaderFooterShapes()
			'ExStart
			'ExFor:Section.DeleteHeaderFooterShapes
			'ExSummary:Removes all images and shapes from all headers footers in a section.
			Dim doc As New Document(MyDir & "Document.doc")
			Dim section As Section = doc.Sections(0)
			section.DeleteHeaderFooterShapes()
			'ExEnd
		End Sub


		<Test> _
		Public Sub SectionsCloneSection()
			'ExStart
			'ExId:SectionsCloneSection
			'ExSummary:Shows how to create a duplicate of a particular section.
			Dim doc As New Document(MyDir & "Document.doc")
			Dim cloneSection As Section = doc.Sections(0).Clone()
			'ExEnd
		End Sub

		<Test> _
		Public Sub SectionsImportSection()
			'ExStart
			'ExId:SectionsImportSection
			'ExSummary:Shows how to copy sections between documents.
			Dim srcDoc As New Document(MyDir & "Document.doc")
			Dim dstDoc As New Document()

			Dim sourceSection As Section = srcDoc.Sections(0)
			Dim newSection As Section = CType(dstDoc.ImportNode(sourceSection, True), Section)
			dstDoc.Sections.Add(newSection)
			'ExEnd
		End Sub

		<Test> _
		Public Sub MigrateFrom2XImportSection()
			Dim srcDoc As New Document()
			Dim dstDoc As New Document()

			'ExStart
			'ExId:MigrateFrom2XImportSection
			'ExSummary:This fragment shows how to insert a section from another document in Aspose.Words 3.0 or higher.
			Dim sourceSection As Section = srcDoc.Sections(0)
			Dim newSection As Section = CType(dstDoc.ImportNode(sourceSection, True), Section)
			dstDoc.Sections.Add(newSection)
			'ExEnd
		End Sub

		<Test> _
		Public Sub ModifyPageSetupInAllSections()
			'ExStart
			'ExId:ModifyPageSetupInAllSections
			'ExSummary:Shows how to set paper size for the whole document.
			Dim doc As New Document(MyDir & "Section.ModifyPageSetupInAllSections.doc")

			' It is important to understand that a document can contain many sections and each
			' section has its own page setup. In this case we want to modify them all.
			For Each section As Section In doc
				section.PageSetup.PaperSize = PaperSize.Letter
			Next section

			doc.Save(MyDir & "\Artifacts\Section.ModifyPageSetupInAllSections.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub CultureInfoPageSetupDefaults()
			Thread.CurrentThread.CurrentCulture = New CultureInfo("en-us")

			Dim docEn As New Document()

			'Assert that page defaults comply current culture info
			Dim sectionEn As Section = docEn.Sections(0)
			Assert.AreEqual(72.0, sectionEn.PageSetup.LeftMargin) ' 2.54 cm
			Assert.AreEqual(72.0, sectionEn.PageSetup.RightMargin) ' 2.54 cm
			Assert.AreEqual(72.0, sectionEn.PageSetup.TopMargin) ' 2.54 cm
			Assert.AreEqual(72.0, sectionEn.PageSetup.BottomMargin) ' 2.54 cm
			Assert.AreEqual(36.0, sectionEn.PageSetup.HeaderDistance) ' 1.27 cm
			Assert.AreEqual(36.0, sectionEn.PageSetup.FooterDistance) ' 1.27 cm
			Assert.AreEqual(36.0, sectionEn.PageSetup.TextColumns.Spacing) ' 1.27 cm

			'Change culture and assert that the page defaults are changed
			Thread.CurrentThread.CurrentCulture = New CultureInfo("de-de")

			Dim docDe As New Document()

			Dim sectionDe As Section = docDe.Sections(0)
			Assert.AreEqual(70.85, sectionDe.PageSetup.LeftMargin) ' 2.5 cm
			Assert.AreEqual(70.85, sectionDe.PageSetup.RightMargin) ' 2.5 cm
			Assert.AreEqual(70.85, sectionDe.PageSetup.TopMargin) ' 2.5 cm
			Assert.AreEqual(56.7, sectionDe.PageSetup.BottomMargin) ' 2 cm
			Assert.AreEqual(35.4, sectionDe.PageSetup.HeaderDistance) ' 1.25 cm
			Assert.AreEqual(35.4, sectionDe.PageSetup.FooterDistance) ' 1.25 cm
			Assert.AreEqual(35.4, sectionDe.PageSetup.TextColumns.Spacing) ' 1.25 cm

			'Change page defaults
			sectionDe.PageSetup.LeftMargin = 90 ' 3.17 cm
			sectionDe.PageSetup.RightMargin = 90 ' 3.17 cm
			sectionDe.PageSetup.TopMargin = 72 ' 2.54 cm
			sectionDe.PageSetup.BottomMargin = 72 ' 2.54 cm
			sectionDe.PageSetup.HeaderDistance = 35.4 ' 1.25 cm
			sectionDe.PageSetup.FooterDistance = 35.4 ' 1.25 cm
			sectionDe.PageSetup.TextColumns.Spacing = 35.4 ' 1.25 cm

			Dim dstStream As New MemoryStream()
			docDe.Save(dstStream, SaveFormat.Docx)

			Dim sectionDeAfter As Section = docDe.Sections(0)
			Assert.AreEqual(90.0, sectionDeAfter.PageSetup.LeftMargin) ' 3.17 cm
			Assert.AreEqual(90.0, sectionDeAfter.PageSetup.RightMargin) ' 3.17 cm
			Assert.AreEqual(72.0, sectionDeAfter.PageSetup.TopMargin) ' 2.54 cm
			Assert.AreEqual(72.0, sectionDeAfter.PageSetup.BottomMargin) ' 2.54 cm
			Assert.AreEqual(35.4, sectionDeAfter.PageSetup.HeaderDistance) ' 1.25 cm
			Assert.AreEqual(35.4, sectionDeAfter.PageSetup.FooterDistance) ' 1.25 cm
			Assert.AreEqual(35.4, sectionDeAfter.PageSetup.TextColumns.Spacing) ' 1.25 cm
		End Sub
	End Class
End Namespace
