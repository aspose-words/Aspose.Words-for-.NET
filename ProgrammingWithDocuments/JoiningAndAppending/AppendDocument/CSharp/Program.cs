//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Reflection;
using System.Collections;
using System.IO;
using System.Text;

using Aspose.Words.Lists;
using Aspose.Words.Fields;
using Aspose.Words;

namespace AppendDocumentExample
{
    public class Program
    {
        private static string gDataDir;

        public static void Main()
        {
            // The path to the documents directory.
            gDataDir = new Uri(new Uri(exeDir), @"../../Data/").LocalPath;

            // Run each of the sample code snippets.
            AppendDocument_SimpleAppendDocument();
            AppendDocument_KeepSourceFormatting();
            AppendDocument_UseDestinationStyles();
            AppendDocument_JoinContinuous();
            AppendDocument_JoinNewPage();
            AppendDocument_RestartPageNumbering();
            AppendDocument_LinkHeadersFooters();
            AppendDocument_UnlinkHeadersFooters();
            AppendDocument_RemoveSourceHeadersFooters();
            AppendDocument_DifferentPageSetup();
            AppendDocument_ConvertNumPageFields();
            AppendDocument_ListUseDestinationStyles();
            AppendDocument_ListKeepSourceFormatting();
            AppendDocument_KeepSourceTogether();
            AppendDocument_BaseDocument();
            AppendDocument_UpdatePageLayout();
        }

        public static void AppendDocument_SimpleAppendDocument()
        {
            Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
            Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

            //ExStart
            //ExId:AppendDocument_SimpleAppend
            //ExSummary:Shows how to append a document to the end of another document using no additional options.
            // Append the source document to the destination document using no extra options.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            //ExEnd

            dstDoc.Save(gDataDir + "TestFile.SimpleAppendDocument Out.docx");
        }

        public static void AppendDocument_KeepSourceFormatting()
        {
            //ExStart
            //ExId:AppendDocument_KeepSourceFormatting
            //ExSummary:Shows how to append a document to another document while keeping the original formatting.
            // Load the documents to join.
            Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
            Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

            // Keep the formatting from the source document when appending it to the destination document.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the joined document to disk.
            dstDoc.Save(gDataDir + "TestFile.KeepSourceFormatting Out.docx");
            //ExEnd
        }

        public static void AppendDocument_UseDestinationStyles()
        {
            //ExStart
            //ExId:AppendDocument_UseDestinationStyles
            //ExSummary:Shows how to append a document to another document using the formatting of the destination document.
            // Load the documents to join.
            Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
            Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

            // Append the source document using the styles of the destination document.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

            // Save the joined document to disk.
            dstDoc.Save(gDataDir + "TestFile.UseDestinationStyles Out.doc");
            //ExEnd
        }

        public static void AppendDocument_JoinContinuous()
        {
            //ExStart
            //ExId:AppendDocument_JoinContinuous
            //ExSummary:Shows how to append a document to another document so the content flows continuously.
            Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
            Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

            // Make the document appear straight after the destination documents content.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            // Append the source document using the original styles found in the source document.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            dstDoc.Save(gDataDir + "TestFile.JoinContinuous Out.doc");
            //ExEnd
        }

        public static void AppendDocument_JoinNewPage()
        {
            //ExStart
            //ExId:AppendDocument_JoinNewPage
            //ExSummary:Shows how to append a document to another document so it starts on a new page.
            Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
            Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

            // Set the appended document to start on a new page.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;

            // Append the source document using the original styles found in the source document.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            dstDoc.Save(gDataDir + "TestFile.JoinNewPage Out.doc");
            //ExEnd
        }

        public static void AppendDocument_RestartPageNumbering()
        {
            //ExStart
            //ExId:AppendDocument_RestartPageNumbering
            //ExSummary:Shows how to append a document to another document with page numbering restarted. 
            Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
            Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

            // Set the appended document to appear on the next page.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;
            // Restart the page numbering for the document to be appended.
            srcDoc.FirstSection.PageSetup.RestartPageNumbering = true;

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            dstDoc.Save(gDataDir + "TestFile.RestartPageNumbering Out.doc");
            //ExEnd
        }

        public static void AppendDocument_LinkHeadersFooters()
        {
            //ExStart
            //ExFor:HeaderFooterCollection.LinkToPrevious(Boolean)
            //ExId:AppendDocument_LinkHeadersFooters
            //ExSummary:Shows how to append a document to another document and continue headers and footers from the destination document.
            Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
            Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

            // Set the appended document to appear on a new page.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;

            // Link the headers and footers in the source document to the previous section. 
            // This will override any headers or footers already found in the source document. 
            srcDoc.FirstSection.HeadersFooters.LinkToPrevious(true);

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            dstDoc.Save(gDataDir + "TestFile.LinkHeadersFooters Out.doc");
            //ExEnd
        }

        public static void AppendDocument_UnlinkHeadersFooters()
        {
            //ExStart
            //ExId:AppendDocument_UnlinkHeadersFooters
            //ExSummary:Shows how to append a document to another document so headers and footers do not continue from the destination document.
            Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
            Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

            // Even a document with no headers or footers can still have the LinkToPrevious setting set to true.
            // Unlink the headers and footers in the source document to stop this from continuing the headers and footers
            // from the destination document.
            srcDoc.FirstSection.HeadersFooters.LinkToPrevious(false);

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            dstDoc.Save(gDataDir + "TestFile.UnlinkHeadersFooters Out.doc");
            //ExEnd
        }

        public static void AppendDocument_RemoveSourceHeadersFooters()
        {
            //ExStart
            //ExId:AppendDocument_RemoveSourceHeadersFooters
            //ExSummary:Shows how to remove headers and footers from a document before appending it to another document. 
            Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
            Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

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
            dstDoc.Save(gDataDir + "TestFile.RemoveSourceHeadersFooters Out.doc");
            //ExEnd
        }

        public static void AppendDocument_DifferentPageSetup()
        {
            //ExStart
            //ExId:AppendDocument_DifferentPageSetup
            //ExSummary:Shows how to append a document to another document continuously which has different page settings.
            Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
            Document srcDoc =  new Document(gDataDir + "TestFile.SourcePageSetup.doc");

            // Set the source document to continue straight after the end of the destination document.
            // If some page setup settings are different then this may not work and the source document will appear 
            // on a new page.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            // To ensure this does not happen when the source document has different page setup settings make sure the
            // settings are identical between the last section of the destination document.
            // If there are further continuous sections that follow on in the source document then this will need to be 
            // repeated for those sections as well.
            srcDoc.FirstSection.PageSetup.PageWidth = dstDoc.LastSection.PageSetup.PageWidth;
            srcDoc.FirstSection.PageSetup.PageHeight = dstDoc.LastSection.PageSetup.PageHeight;
            srcDoc.FirstSection.PageSetup.Orientation = dstDoc.LastSection.PageSetup.Orientation;

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            dstDoc.Save(gDataDir + "TestFile.DifferentPageSetup Out.doc");
            //ExEnd
        }

        //ExStart
        //ExId:AppendDocument_ConvertNumPageFields
        //ExSummary:Shows how to change the NUMPAGE fields in a document to display the number of pages only within a sub document.
        public static void AppendDocument_ConvertNumPageFields()
        {
            Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
            Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

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

            dstDoc.Save(gDataDir + "TestFile.ConvertNumPageFields Out.doc");
        }

        /// <summary>
        /// Replaces all NUMPAGES fields in the document with PAGEREF fields. The replacement field displays the total number
        /// of pages in the sub document instead of the total pages in the document.
        /// </summary>
        /// <param name="doc">The combined document to process</param>
        public static void ConvertNumPageFieldsToPageRef(Document doc)
        {
            // This is the prefix for each bookmark which signals where page numbering restarts.
            // The underscore "_" at the start inserts this bookmark as hidden in MS Word.
            const string bookmarkPrefix = "_SubDocumentEnd";
            // Field name of the NUMPAGES field.
            const string numPagesFieldName = "NUMPAGES";
            // Field name of the PAGEREF field.
            const string pageRefFieldName = "PAGEREF";

            // Create a new DocumentBuilder which is used to insert the bookmarks and replacement fields.
            DocumentBuilder builder = new DocumentBuilder(doc);
            // Defines the number of page restarts that have been encountered and therefore the number of "sub" documents
            // found within this document.
            int subDocumentCount = 0;

            // Iterate through all sections in the document.
            foreach (Section section in doc.Sections)
            {
                // This section has it's page numbering restarted so we will treat this as the start of a sub document.
                // Any PAGENUM fields in this inner document must be converted to special PAGEREF fields to correct numbering.
                if (section.PageSetup.RestartPageNumbering)
                {
                    // Don't do anything if this is the first section in the document. This part of the code will insert the bookmark marking
                    // the end of the previous sub document so therefore it is not applicable for first section in the document.
                    if (!section.Equals(doc.FirstSection))
                    {
                        // Get the previous section and the last node within the body of that section.
                        Section prevSection = (Section)section.PreviousSibling;
                        Node lastNode = prevSection.Body.LastChild;

                        // Use the DocumentBuilder to move to this node and insert the bookmark there.
                        // This bookmark represents the end of the sub document.
                        builder.MoveTo(lastNode);
                        builder.StartBookmark(bookmarkPrefix + subDocumentCount);
                        builder.EndBookmark(bookmarkPrefix + subDocumentCount);

                        // Increase the subdocument count to insert the correct bookmarks.
                        subDocumentCount++;
                    }
                }

                // The last section simply needs the ending bookmark to signal that it is the end of the current sub document.
                if (section.Equals(doc.LastSection))
                {
                    // Insert the bookmark at the end of the body of the last section.
                    // Don't increase the count this time as we are just marking the end of the document.
                    Node lastNode = doc.LastSection.Body.LastChild;
                    builder.MoveTo(lastNode);
                    builder.StartBookmark(bookmarkPrefix + subDocumentCount);
                    builder.EndBookmark(bookmarkPrefix + subDocumentCount);
                }

                // Iterate through each NUMPAGES field in the section and replace the field with a PAGEREF field referring to the bookmark of the current subdocument
                // This bookmark is positioned at the end of the sub document but does not exist yet. It is inserted when a section with restart page numbering or the last 
                // section is encountered.
                Node[] nodes = section.GetChildNodes(NodeType.FieldStart, true).ToArray();
                foreach (FieldStart fieldStart in nodes)
                {
                    if (fieldStart.FieldType == FieldType.FieldNumPages)
                    {
                        // Get the field code.
                        string fieldCode = GetFieldCode(fieldStart);
                        // Since the NUMPAGES field does not take any additional parameters we can assume the remaining part of the field
                        // code after the fieldname are the switches. We will use these to help recreate the NUMPAGES field as a PAGEREF field.
                        string fieldSwitches = fieldCode.Replace(numPagesFieldName, "").Trim();

                        // Inserting the new field directly at the FieldStart node of the original field will cause the new field to
                        // not pick up the formatting of the original field. To counter this insert the field just before the original field
                        Node previousNode = fieldStart.PreviousSibling;

                        // If a previous run cannot be found then we are forced to use the FieldStart node.
                        if (previousNode == null)
                            previousNode = fieldStart;

                        // Insert a PAGEREF field at the same position as the field.
                        builder.MoveTo(previousNode);
                        // This will insert a new field with a code like " PAGEREF _SubDocumentEnd0 *\MERGEFORMAT ".
                        Field newField = builder.InsertField(string.Format(" {0} {1}{2} {3} ", pageRefFieldName, bookmarkPrefix, subDocumentCount, fieldSwitches));

                        // The field will be inserted before the referenced node. Move the node before the field instead.
                        previousNode.ParentNode.InsertBefore(previousNode, newField.Start);

                        // Remove the original NUMPAGES field from the document.
                        RemoveField(fieldStart);
                    }
                }
            }
        }
        //ExEnd

        public static void AppendDocument_ListUseDestinationStyles()
        {
            //ExStart
            //ExId:AppendDocument_ListUseDestinationStyles
            //ExSummary:Shows how to append a document using destination styles and preventing any list numberings from continuing on.
            Document dstDoc = new Document(gDataDir + "TestFile.DestinationList.doc");
            Document srcDoc =  new Document(gDataDir + "TestFile.SourceList.doc");

            // Set the source document to continue straight after the end of the destination document.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            // Keep track of the lists that are created.
            Hashtable newLists = new Hashtable();

            // Iterate through all paragraphs in the document.
            foreach (Paragraph para in srcDoc.GetChildNodes(NodeType.Paragraph, true))
            {
                if (para.IsListItem)
                {
                    int listId = para.ListFormat.List.ListId;

                    // Check if the destination document contains a list with this ID already. If it does then this may
                    // cause the two lists to run together. Create a copy of the list in the source document instead.
                    if (dstDoc.Lists.GetListByListId(listId) != null)
                    {
                        List currentList;
                        // A newly copied list already exists for this ID, retrieve the stored list and use it on 
                        // the current paragraph.
                        if (newLists.Contains(listId))
                        {
                            currentList = (List)newLists[listId];
                        }
                        else
                        {
                            // Add a copy of this list to the document and store it for later reference.
                            currentList = srcDoc.Lists.AddCopy(para.ListFormat.List);
                            newLists.Add(listId, currentList);
                        }

                        // Set the list of this paragraph  to the copied list.
                        para.ListFormat.List = currentList;
                    }
                }
            }

            // Append the source document to end of the destination document.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

            // Save the combined document to disk.
            dstDoc.Save(gDataDir + "TestFile.ListUseDestinationStyles Out.docx");
            //ExEnd
        }

        public static void AppendDocument_ListKeepSourceFormatting()
        {
            //ExStart
            //ExId:AppendDocument_ListKeepSourceFormatting
            //ExSummary:Shows how to append a document to another document containing lists retaining source formatting.
            Document dstDoc = new Document(gDataDir + "TestFile.DestinationList.doc");
            Document srcDoc =  new Document(gDataDir + "TestFile.SourceList.doc");

            // Append the content of the document so it flows continuously.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            dstDoc.Save(gDataDir + "TestFile.ListKeepSourceFormatting Out.doc");
            //ExEnd
        }

        public static void AppendDocument_KeepSourceTogether()
        {
            //ExStart
            //ExFor:ParagraphFormat.KeepWithNext
            //ExId:AppendDocument_KeepSourceTogether
            //ExSummary:Shows how to append a document to another document while keeping the content from splitting across two pages.
            Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
            Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

            // Set the source document to appear straight after the destination document's content.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            // Iterate through all sections in the source document.
            foreach(Paragraph para in srcDoc.GetChildNodes(NodeType.Paragraph, true))
            {
                para.ParagraphFormat.KeepWithNext = true;
            }

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            dstDoc.Save(gDataDir + "TestDcc.KeepSourceTogether Out.doc");
            //ExEnd
        }

        public static void AppendDocument_BaseDocument()
        {
            //ExStart
            //ExId:AppendDocument_BaseDocument
            //ExSummary:Shows how to remove all content from a document before using it as a base to append documents to.
            // Use a blank document as the destination document.
            Document dstDoc = new Document();
            Document srcDoc = new Document(gDataDir + "TestFile.Source.doc");

            // The destination document is not actually empty which often causes a blank page to appear before the appended document
            // This is due to the base document having an empty section and the new document being started on the next page.
            // Remove all content from the destination document before appending.
            dstDoc.RemoveAllChildren();

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            dstDoc.Save(gDataDir + "TestFile.BaseDocument Out.doc");
            //ExEnd
        }

        public static void AppendDocument_UpdatePageLayout()
        {
            //ExStart
            //ExId:AppendDocument_UpdatePageLayout
            //ExSummary:Shows how to rebuild the document layout after appending further content.
            Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
            Document srcDoc = new Document(gDataDir + "TestFile.Source.doc");

            // If the destination document is rendered to PDF, image etc or UpdatePageLayout is called before the source document 
            // is appended then any changes made after will not be reflected in the rendered output.
            dstDoc.UpdatePageLayout();

            // Join the documents.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // For the changes to be updated to rendered output, UpdatePageLayout must be called again.
            // If not called again the appended document will not appear in the output of the next rendering.
            dstDoc.UpdatePageLayout();

            // Save the joined document to PDF.
            dstDoc.Save(gDataDir + "TestFile.UpdatePageLayout Out.pdf");
            //ExEnd
        }

        //ExStart
        //ExFor:FieldStart
        //ExFor:FieldSeparator
        //ExFor:FieldEnd
        //ExId:AppendDocument_HelperFunctions
        //ExSummary:Provides some helper functions by the methods above
        /// <summary>
        /// Retrieves the field code from a field.
        /// </summary>
        /// <param name="fieldStart">The field start of the field which to gather the field code from</param>
        /// <returns></returns>
        private static string GetFieldCode(FieldStart fieldStart)
        {
            StringBuilder builder = new StringBuilder();

            for (Node node = fieldStart; node != null && node.NodeType != NodeType.FieldSeparator && 
                node.NodeType != NodeType.FieldEnd; node = node.NextPreOrder(node.Document))
            {
                // Use text only of Run nodes to avoid duplication.
                if (node.NodeType == NodeType.Run)
                    builder.Append(node.GetText());
            }
            return builder.ToString();
        }

        /// <summary>
        /// Removes the Field from the document
        /// </summary>
        /// <param name="fieldStart">The field start node of the field to remove.</param>
        private static void RemoveField(FieldStart fieldStart)
        {
            Node currentNode = fieldStart;
            bool isRemoving = true;
            while (currentNode != null && isRemoving)
            {
                if (currentNode.NodeType == NodeType.FieldEnd)
                    isRemoving = false;

                Node nextNode = currentNode.NextPreOrder(currentNode.Document);
                currentNode.Remove();
                currentNode = nextNode;
            }
        }
        //ExEnd

        //ExStart
        //ExFor:DocumentBase.ImportNode(Node,bool,ImportFormatMode)
        //ExFor:ImportFormatMode
        //ExId:CombineDocuments
        //ExSummary:Shows how to manually append the content from one document to the end of another document.
        /// <summary>
        /// A manual implementation of the Document.AppendDocument function which shows the general 
        /// steps of how a document is appended to another.
        /// </summary>
        /// <param name="dstDoc">The destination document where to append to.</param>
        /// <param name="srcDoc">The source document.</param>
        /// <param name="mode">The import mode to use when importing content from another document.</param>
        public void AppendDocument(Document dstDoc, Document srcDoc, ImportFormatMode mode)
        {
            // Loop through all sections in the source document. 
            // Section nodes are immediate children of the Document node so we can just enumerate the Document.
            foreach (Section srcSection in srcDoc)
            {
                // Because we are copying a section from one document to another, 
                // it is required to import the Section node into the destination document.
                // This adjusts any document-specific references to styles, lists, etc.
                //
                // Importing a node creates a copy of the original node, but the copy
                // is ready to be inserted into the destination document.
                Node dstSection = dstDoc.ImportNode(srcSection, true, mode);

                // Now the new section node can be appended to the destination document.
                dstDoc.AppendChild(dstSection);
            }
        }
        //ExEnd

        //ExStart
        //ExFor:DocumentBase.ImportNode(Node,bool,ImportFormatMode)
        //ExFor:CompositeNode.PrependChild(Node)
        //ExFor:ImportFormatMode
        //ExId:PrependDocument
        //ExSummary:Shows how to manually prepend the content from one document to the beginning of another document.
        public static void PrependDocumentMain()
        {
            Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
            Document srcDoc = new Document(gDataDir + "TestFile.Source.doc");

            // Append the source document to the destination document. This causes the result to have line spacing problems.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Instead prepend the content of the destination document to the start of the source document.
            // This results in the same joined document but with no line spacing issues.
            PrependDocument(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        }

      
        /// <summary>
        /// A modified version of the AppendDocument method which prepends the content of one document to the start
        /// of another.
        /// </summary>
        /// <param name="dstDoc">The destination document where to prepend the source document to.</param>
        /// <param name="srcDoc">The source document.</param>
        public static void PrependDocument(Document dstDoc, Document srcDoc, ImportFormatMode mode)
        {
            // Loop through all sections in the source document. 
            // Section nodes are immediate children of the Document node so we can just enumerate the Document.
            ArrayList sections = new ArrayList(srcDoc.Sections.ToArray());

            // Reverse the order of the sections so they are prepended to start of the destination document in the correct order.
            sections.Reverse();

            foreach (Section srcSection in sections)
            {
                // Import the nodes from the source document.
                Node dstSection = dstDoc.ImportNode(srcSection, true, mode);

                // Now the new section node can be prepended to the destination document.
                // Note how PrependChild is used instead of AppendChild. This is the only line changed compared 
                // to the original method.
                dstDoc.PrependChild(dstSection);
            }
        }
        //ExEnd
          
    }
}