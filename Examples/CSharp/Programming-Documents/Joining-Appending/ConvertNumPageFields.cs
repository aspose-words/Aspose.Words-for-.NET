using System;
using System.IO;

using Aspose.Words;
using Aspose.Words.Fields;
using System.Text;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Joining_and_Appending
{
    class ConvertNumPageFields
    {
        public static void Run()
        {
            //ExStart:ConvertNumPageFields
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_JoiningAndAppending();
            string fileName = "TestFile.Destination.doc";

            Document dstDoc = new Document(dataDir + fileName);
            Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

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

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            dstDoc.Save(dataDir);
            //ExEnd:ConvertNumPageFields
            Console.WriteLine("\nDocument appended successfully with conversion of NUMPAGE fields with PAGEREF fields.\nFile saved at " + dataDir);
        }
        //ExStart:ConvertNumPageFieldsToPageRef
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
        //ExEnd:ConvertNumPageFieldsToPageRef
        //ExStart:GetRemoveField
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
        //ExEnd:GetRemoveField
    }
}
