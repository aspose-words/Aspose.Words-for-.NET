using System.Collections;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Working_with_Document
{
    internal class JoinAndAppendDocuments : DocsExamplesBase
    {
        [Test]
        public void SimpleAppendDocument()
        {
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Append the source document to the destination document using no extra options.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            dstDoc.Save(ArtifactsDir + "JoinAndAppendDocuments.SimpleAppendDocument.docx");
        }

        [Test]
        public void AppendDocument()
        {
            //ExStart:AppendDocumentManually
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");
            
            // Loop through all sections in the source document.
            // Section nodes are immediate children of the Document node so we can just enumerate the Document.
            foreach (Section srcSection in srcDoc)
            {
                // Because we are copying a section from one document to another, 
                // it is required to import the Section node into the destination document.
                // This adjusts any document-specific references to styles, lists, etc.
                //
                // Importing a node creates a copy of the original node, but the copy
                // ss ready to be inserted into the destination document.
                Node dstSection = dstDoc.ImportNode(srcSection, true, ImportFormatMode.KeepSourceFormatting);

                // Now the new section node can be appended to the destination document.
                dstDoc.AppendChild(dstSection);
            }

            dstDoc.Save(ArtifactsDir + "JoinAndAppendDocuments.AppendDocument.docx");
            //ExEnd:AppendDocumentManually
        }

        [Test]
        public void AppendDocumentToBlank()
        {
            //ExStart:AppendDocumentToBlank
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document();
            
            // The destination document is not empty, often causing a blank page to appear before the appended document.
            // This is due to the base document having an empty section and the new document being started on the next page.
            // Remove all content from the destination document before appending.
            dstDoc.RemoveAllChildren();
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "JoinAndAppendDocuments.AppendDocumentToBlank.docx");
            //ExEnd:AppendDocumentToBlank
        }

        [Test]
        public void AppendWithImportFormatOptions()
        {
            //ExStart:AppendWithImportFormatOptions
            Document srcDoc = new Document(MyDir + "Document source with list.docx");
            Document dstDoc = new Document(MyDir + "Document destination with list.docx");

            // Specify that if numbering clashes in source and destination documents,
            // then numbering from the source document will be used.
            ImportFormatOptions options = new ImportFormatOptions { KeepSourceNumbering = true };
            
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles, options);
            //ExEnd:AppendWithImportFormatOptions
        }

        [Test]
        public void ConvertNumPageFields()
        {
            //ExStart:ConvertNumPageFields
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Restart the page numbering on the start of the source document.
            srcDoc.FirstSection.PageSetup.RestartPageNumbering = true;
            srcDoc.FirstSection.PageSetup.PageStartingNumber = 1;

            // Append the source document to the end of the destination document.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // After joining the documents the NUMPAGE fields will now display the total number of pages which
            // is undesired behavior. Call this method to fix them by replacing them with PAGEREF fields.
            ConvertNumPageFieldsToPageRef(dstDoc);

            // This needs to be called in order to update the new fields with page numbers.
            dstDoc.UpdatePageLayout();

            dstDoc.Save(ArtifactsDir + "JoinAndAppendDocuments.ConvertNumPageFields.docx");
            //ExEnd:ConvertNumPageFields
        }

        //ExStart:ConvertNumPageFieldsToPageRef
        public void ConvertNumPageFieldsToPageRef(Document doc)
        {
            // This is the prefix for each bookmark, which signals where page numbering restarts.
            // The underscore "_" at the start inserts this bookmark as hidden in MS Word.
            const string bookmarkPrefix = "_SubDocumentEnd";
            const string numPagesFieldName = "NUMPAGES";
            const string pageRefFieldName = "PAGEREF";

            // Defines the number of page restarts encountered and, therefore,
            // the number of "sub" documents found within this document.
            int subDocumentCount = 0;

            DocumentBuilder builder = new DocumentBuilder(doc);
            
            foreach (Section section in doc.Sections)
            {
                // This section has its page numbering restarted to treat this as the start of a sub-document.
                // Any PAGENUM fields in this inner document must be converted to special PAGEREF fields to correct numbering.
                if (section.PageSetup.RestartPageNumbering)
                {
                    // Don't do anything if this is the first section of the document.
                    // This part of the code will insert the bookmark marking the end of the previous sub-document so,
                    // therefore, it does not apply to the first section in the document.
                    if (!section.Equals(doc.FirstSection))
                    {
                        // Get the previous section and the last node within the body of that section.
                        Section prevSection = (Section) section.PreviousSibling;
                        Node lastNode = prevSection.Body.LastChild;

                        builder.MoveTo(lastNode);
                        
                        // This bookmark represents the end of the sub-document.
                        builder.StartBookmark(bookmarkPrefix + subDocumentCount);
                        builder.EndBookmark(bookmarkPrefix + subDocumentCount);

                        // Increase the sub-document count to insert the correct bookmarks.
                        subDocumentCount++;
                    }
                }

                // The last section needs the ending bookmark to signal that it is the end of the current sub-document.
                if (section.Equals(doc.LastSection))
                {
                    // Insert the bookmark at the end of the body of the last section.
                    // Don't increase the count this time as we are just marking the end of the document.
                    Node lastNode = doc.LastSection.Body.LastChild;
                    
                    builder.MoveTo(lastNode);
                    builder.StartBookmark(bookmarkPrefix + subDocumentCount);
                    builder.EndBookmark(bookmarkPrefix + subDocumentCount);
                }

                // Iterate through each NUMPAGES field in the section and replace it with a PAGEREF field
                // referring to the bookmark of the current sub-document. This bookmark is positioned at the end
                // of the sub-document but does not exist yet. It is inserted when a section with restart page numbering
                // or the last section is encountered.
                Node[] nodes = section.GetChildNodes(NodeType.FieldStart, true).ToArray();
                
                foreach (FieldStart fieldStart in nodes)
                {
                    if (fieldStart.FieldType == FieldType.FieldNumPages)
                    {
                        string fieldCode = GetFieldCode(fieldStart);
                        // Since the NUMPAGES field does not take any additional parameters,
                        // we can assume the field's remaining part. Code after the field name is the switches.
                        // We will use these to help recreate the NUMPAGES field as a PAGEREF field.
                        string fieldSwitches = fieldCode.Replace(numPagesFieldName, "").Trim();

                        // Inserting the new field directly at the FieldStart node of the original field will cause
                        // the new field not to pick up the original field's formatting. To counter this,
                        // insert the field just before the original field if a previous run cannot be found,
                        // we are forced to use the FieldStart node.
                        Node previousNode = fieldStart.PreviousSibling ?? fieldStart;
                        
                        // Insert a PAGEREF field at the same position as the field.
                        builder.MoveTo(previousNode);
                        
                        Field newField = builder.InsertField(
                            $" {pageRefFieldName} {bookmarkPrefix}{subDocumentCount} {fieldSwitches} ");

                        // The field will be inserted before the referenced node. Move the node before the field instead.
                        previousNode.ParentNode.InsertBefore(previousNode, newField.Start);

                        // Remove the original NUMPAGES field from the document.
                        RemoveField(fieldStart);
                    }
                }
            }
        }
        //ExEnd:ConvertNumPageFieldsToPageRef
        
        //ExStart:GetRemoveField
        private void RemoveField(FieldStart fieldStart)
        {
            bool isRemoving = true;
            
            Node currentNode = fieldStart;
            while (currentNode != null && isRemoving)
            {
                if (currentNode.NodeType == NodeType.FieldEnd)
                    isRemoving = false;

                Node nextNode = currentNode.NextPreOrder(currentNode.Document);
                currentNode.Remove();
                currentNode = nextNode;
            }
        }

        private string GetFieldCode(FieldStart fieldStart)
        {
            StringBuilder builder = new StringBuilder();

            for (Node node = fieldStart;
                node != null && node.NodeType != NodeType.FieldSeparator &&
                node.NodeType != NodeType.FieldEnd;
                node = node.NextPreOrder(node.Document))
            {
                // Use text only of Run nodes to avoid duplication.
                if (node.NodeType == NodeType.Run)
                    builder.Append(node.GetText());
            }

            return builder.ToString();
        }
        //ExEnd:GetRemoveField

        [Test]
        public void DifferentPageSetup()
        {
            //ExStart:DifferentPageSetup
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Set the source document to continue straight after the end of the destination document.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            // Restart the page numbering on the start of the source document.
            srcDoc.FirstSection.PageSetup.RestartPageNumbering = true;
            srcDoc.FirstSection.PageSetup.PageStartingNumber = 1;

            // To ensure this does not happen when the source document has different page setup settings, make sure the
            // settings are identical between the last section of the destination document.
            // If there are further continuous sections that follow on in the source document,
            // this will need to be repeated for those sections.
            srcDoc.FirstSection.PageSetup.PageWidth = dstDoc.LastSection.PageSetup.PageWidth;
            srcDoc.FirstSection.PageSetup.PageHeight = dstDoc.LastSection.PageSetup.PageHeight;
            srcDoc.FirstSection.PageSetup.Orientation = dstDoc.LastSection.PageSetup.Orientation;

            // Iterate through all sections in the source document.
            foreach (Paragraph para in srcDoc.GetChildNodes(NodeType.Paragraph, true))
            {
                para.ParagraphFormat.KeepWithNext = true;
            }

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "JoinAndAppendDocuments.DifferentPageSetup.docx");
            //ExEnd:DifferentPageSetup
        }

        [Test]
        public void JoinContinuous()
        {
            //ExStart:JoinContinuous
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Make the document appear straight after the destination documents content.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;
            // Append the source document using the original styles found in the source document.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "JoinAndAppendDocuments.JoinContinuous.docx");
            //ExEnd:JoinContinuous
        }

        [Test]
        public void JoinNewPage()
        {
            //ExStart:JoinNewPage
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Set the appended document to start on a new page.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;
            // Append the source document using the original styles found in the source document.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "JoinAndAppendDocuments.JoinNewPage.docx");
            //ExEnd:JoinNewPage
        }

        [Test]
        public void KeepSourceFormatting()
        {
            //ExStart:KeepSourceFormatting
            Document dstDoc = new Document();
            dstDoc.FirstSection.Body.AppendParagraph("Destination document text. ");

            Document srcDoc = new Document();
            srcDoc.FirstSection.Body.AppendParagraph("Source document text. ");

            // Append the source document to the destination document.
            // Pass format mode to retain the original formatting of the source document when importing it.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            dstDoc.Save(ArtifactsDir + "JoinAndAppendDocuments.KeepSourceFormatting.docx");
            //ExEnd:KeepSourceFormatting
        }

        [Test]
        public void KeepSourceTogether()
        {
            //ExStart:KeepSourceTogether
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Document destination with list.docx");
            
            // Set the source document to appear straight after the destination document's content.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            foreach (Paragraph para in srcDoc.GetChildNodes(NodeType.Paragraph, true))
            {
                para.ParagraphFormat.KeepWithNext = true;
            }

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "JoinAndAppendDocuments.KeepSourceTogether.docx");
            //ExEnd:KeepSourceTogether
        }        

        [Test]
        public void ListKeepSourceFormatting()
        {
            //ExStart:ListKeepSourceFormatting
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Document destination with list.docx");

            // Append the content of the document so it flows continuously.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "JoinAndAppendDocuments.ListKeepSourceFormatting.docx");
            //ExEnd:ListKeepSourceFormatting
        }

        [Test]
        public void ListUseDestinationStyles()
        {
            //ExStart:ListUseDestinationStyles
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Document destination with list.docx");

            // Set the source document to continue straight after the end of the destination document.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            // Keep track of the lists that are created.
            Dictionary<int, Aspose.Words.Lists.List> newLists = new Dictionary<int, Aspose.Words.Lists.List>();

            foreach (Paragraph para in srcDoc.GetChildNodes(NodeType.Paragraph, true))
            {
                if (para.IsListItem)
                {
                    int listId = para.ListFormat.List.ListId;

                    // Check if the destination document contains a list with this ID already. If it does, then this may
                    // cause the two lists to run together. Create a copy of the list in the source document instead.
                    if (dstDoc.Lists.GetListByListId(listId) != null)
                    {
                        Aspose.Words.Lists.List currentList;
                        // A newly copied list already exists for this ID, retrieve the stored list,
                        // and use it on the current paragraph.
                        if (newLists.ContainsKey(listId))
                        {
                            currentList = newLists[listId];
                        }
                        else
                        {
                            // Add a copy of this list to the document and store it for later reference.
                            currentList = srcDoc.Lists.AddCopy(para.ListFormat.List);
                            newLists.Add(listId, currentList);
                        }

                        // Set the list of this paragraph to the copied list.
                        para.ListFormat.List = currentList;
                    }
                }
            }

            // Append the source document to end of the destination document.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

            dstDoc.Save(ArtifactsDir + "JoinAndAppendDocuments.ListUseDestinationStyles.docx");
            //ExEnd:ListUseDestinationStyles
        }

        [Test]
        public void RestartPageNumbering()
        {
            //ExStart:RestartPageNumbering
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;
            srcDoc.FirstSection.PageSetup.RestartPageNumbering = true;

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "JoinAndAppendDocuments.RestartPageNumbering.docx");
            //ExEnd:RestartPageNumbering
        }

        [Test]
        public void UpdatePageLayout()
        {
            //ExStart:UpdatePageLayout
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // If the destination document is rendered to PDF, image etc.
            // or UpdatePageLayout is called before the source document. Is appended,
            // then any changes made after will not be reflected in the rendered output
            dstDoc.UpdatePageLayout();

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // For the changes to be updated to rendered output, UpdatePageLayout must be called again.
            // If not called again, the appended document will not appear in the output of the next rendering.
            dstDoc.UpdatePageLayout();

            dstDoc.Save(ArtifactsDir + "JoinAndAppendDocuments.UpdatePageLayout.docx");
            //ExEnd:UpdatePageLayout
        }

        [Test]
        public void UseDestinationStyles()
        {
            //ExStart:UseDestinationStyles
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Append the source document using the styles of the destination document.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

            dstDoc.Save(ArtifactsDir + "JoinAndAppendDocuments.UseDestinationStyles.docx");
            //ExEnd:UseDestinationStyles
        }

        [Test]
        public void SmartStyleBehavior()
        {
            //ExStart:SmartStyleBehavior
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");
            DocumentBuilder builder = new DocumentBuilder(dstDoc);
            
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.PageBreak);

            ImportFormatOptions options = new ImportFormatOptions { SmartStyleBehavior = true };

            builder.InsertDocument(srcDoc, ImportFormatMode.UseDestinationStyles, options);
            builder.Document.Save(ArtifactsDir + "JoinAndAppendDocuments.SmartStyleBehavior.docx");
            //ExEnd:SmartStyleBehavior
        }

        [Test]
        public void InsertDocumentWithBuilder()
        {
            //ExStart:InsertDocumentWithBuilder
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");
            DocumentBuilder builder = new DocumentBuilder(dstDoc);

            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.PageBreak);

            builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            builder.Document.Save(ArtifactsDir + "JoinAndAppendDocuments.InsertDocumentWithBuilder.docx");
            //ExEnd:InsertDocumentWithBuilder
        }

        [Test]
        public void KeepSourceNumbering()
        {
            //ExStart:KeepSourceNumbering
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Keep source list formatting when importing numbered paragraphs.
            ImportFormatOptions importFormatOptions = new ImportFormatOptions { KeepSourceNumbering = true };
            
            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting,
                importFormatOptions);

            ParagraphCollection srcParas = srcDoc.FirstSection.Body.Paragraphs;
            foreach (Paragraph srcPara in srcParas)
            {
                Node importedNode = importer.ImportNode(srcPara, false);
                dstDoc.FirstSection.Body.AppendChild(importedNode);
            }

            dstDoc.Save(ArtifactsDir + "JoinAndAppendDocuments.KeepSourceNumbering.docx");
            //ExEnd:KeepSourceNumbering
        }

        [Test]
        public void IgnoreTextBoxes()
        {
            //ExStart:IgnoreTextBoxes
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Keep the source text boxes formatting when importing.
            ImportFormatOptions importFormatOptions = new ImportFormatOptions { IgnoreTextBoxes = false };
            
            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting,
                importFormatOptions);

            ParagraphCollection srcParas = srcDoc.FirstSection.Body.Paragraphs;
            foreach (Paragraph srcPara in srcParas)
            {
                Node importedNode = importer.ImportNode(srcPara, true);
                dstDoc.FirstSection.Body.AppendChild(importedNode);
            }

            dstDoc.Save(ArtifactsDir + "JoinAndAppendDocuments.IgnoreTextBoxes.docx");
            //ExEnd:IgnoreTextBoxes
        }

        [Test]
        public void IgnoreHeaderFooter()
        {
            //ExStart:IgnoreHeaderFooter
            Document srcDocument = new Document(MyDir + "Document source.docx");
            Document dstDocument = new Document(MyDir + "Northwind traders.docx");

            ImportFormatOptions importFormatOptions = new ImportFormatOptions { IgnoreHeaderFooter = false };

            dstDocument.AppendDocument(srcDocument, ImportFormatMode.KeepSourceFormatting, importFormatOptions);
            
            dstDocument.Save(ArtifactsDir + "JoinAndAppendDocuments.IgnoreHeaderFooter.docx");
            //ExEnd:IgnoreHeaderFooter
        }

        [Test]
        public void LinkHeadersFooters()
        {
            //ExStart:LinkHeadersFooters
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Set the appended document to appear on a new page.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;
            // Link the headers and footers in the source document to the previous section.
            // This will override any headers or footers already found in the source document.
            srcDoc.FirstSection.HeadersFooters.LinkToPrevious(true);

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            dstDoc.Save(ArtifactsDir + "JoinAndAppendDocuments.LinkHeadersFooters.docx");
            //ExEnd:LinkHeadersFooters
        }

        [Test]
        public void RemoveSourceHeadersFooters()
        {
            //ExStart:RemoveSourceHeadersFooters
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Remove the headers and footers from each of the sections in the source document.
            foreach (Section section in srcDoc.Sections)
            {
                section.ClearHeadersFooters();
            }

            // Even after the headers and footers are cleared from the source document, the "LinkToPrevious" setting 
            // for HeadersFooters can still be set. This will cause the headers and footers to continue from the destination 
            // document. This should set to false to avoid this behavior.
            srcDoc.FirstSection.HeadersFooters.LinkToPrevious(false);

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            dstDoc.Save(ArtifactsDir + "JoinAndAppendDocuments.RemoveSourceHeadersFooters.docx");
            //ExEnd:RemoveSourceHeadersFooters
        }

        [Test]
        public void UnlinkHeadersFooters()
        {
            //ExStart:UnlinkHeadersFooters
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Unlink the headers and footers in the source document to stop this
            // from continuing the destination document's headers and footers.
            srcDoc.FirstSection.HeadersFooters.LinkToPrevious(false);

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            dstDoc.Save(ArtifactsDir + "JoinAndAppendDocuments.UnlinkHeadersFooters.docx");
            //ExEnd:UnlinkHeadersFooters
        }
    }
}