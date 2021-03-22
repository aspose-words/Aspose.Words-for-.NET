using System.Collections;
using System.Collections.Generic;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace DocsExamples.Mail_Merge_and_Reporting.Custom_examples
{
    internal class ApplyCustomLogicToEmptyRegions : DocsExamplesBase
    {
        [Test]
        public void ExecuteWithRegionsNestedCustom()
        {
            //ExStart:ApplyCustomLogicToEmptyRegions
            Document doc = new Document(MyDir + "Mail merge destination - Northwind suppliers.docx");

            // Create a data source which has some data missing.
            // This will result in some regions are merged, and some remain after executing mail merge
            DataSet data = GetDataSource();

            // Ensure that we have not set the removal of any unused regions as we will handle them manually.
            // We achieve this by removing the RemoveUnusedRegions flag from the cleanup options using the AND and NOT bitwise operators.
            doc.MailMerge.CleanupOptions = doc.MailMerge.CleanupOptions & ~MailMergeCleanupOptions.RemoveUnusedRegions;

            doc.MailMerge.ExecuteWithRegions(data);

            // Regions without data and not merged will remain in the document.
            Document mergedDoc = doc.Clone(); //ExSkip
            
            // Apply logic to each unused region left in the document.
            ExecuteCustomLogicOnEmptyRegions(doc, new EmptyRegionsHandler());

            doc.Save(ArtifactsDir + "MailMerge.ExecuteWithRegionsNestedCustom_1.docx");

            doc = mergedDoc.Clone();

            // Apply different logic to unused regions this time.
            ExecuteCustomLogicOnEmptyRegions(doc, new EmptyRegionsHandlerMergeTable());

            doc.Save(ArtifactsDir + "MailMerge.ExecuteWithRegionsNestedCustom_2.docx");
            //ExEnd:ApplyCustomLogicToEmptyRegions
            
            doc = mergedDoc.Clone();
            
            //ExStart:ContactDetails 
            List<string> regions = new List<string>();
            regions.Add("ContactDetails");

            // Only handle the ContactDetails region in our handler.
            ExecuteCustomLogicOnEmptyRegions(doc, new EmptyRegionsHandler(), regions);
            //ExEnd:ContactDetails

            doc.Save(ArtifactsDir + "MailMerge.ExecuteWithRegionsNestedCustom_3.docx");
        }

        //ExStart:CreateDataSourceFromDocumentRegions
        /// <summary>
        /// Returns a DataSet object containing a DataTable for the unmerged regions in the specified document.
        /// If regionsList is null all regions found within the document are included. If an List instance is present,
        /// the only regions specified in the list found in the document are added.
        /// </summary>
        private DataSet CreateDataSourceFromDocumentRegions(Document doc, List<string> regionsList)
        {
            const string tableStartMarker = "TableStart:";
            DataSet dataSet = new DataSet();
            string tableName = null;

            foreach (string fieldName in doc.MailMerge.GetFieldNames())
            {
                if (fieldName.Contains(tableStartMarker))
                {
                    tableName = fieldName.Substring(tableStartMarker.Length);
                }
                else if (tableName != null)
                {
                    // Add the table name as a new DataTable if it doesn't already exist in the DataSet.
                    if (dataSet.Tables[tableName] == null)
                    {
                        DataTable table = new DataTable(tableName);
                        table.Columns.Add(fieldName);

                        // We only need to add the first field for the handler to be called for the region's fields.
                        if (regionsList == null || regionsList.Contains(tableName))
                        {
                            table.Rows.Add("FirstField");
                        }

                        dataSet.Tables.Add(table);
                    }

                    tableName = null;
                }
            }

            return dataSet;
        }
        //ExEnd:CreateDataSourceFromDocumentRegions

        //ExStart:ExecuteCustomLogicOnEmptyRegions
        /// <summary>
        /// Applies logic defined in the passed handler class to all unused regions in the document.
        /// This allows controlling how unused regions are handled in the document manually.
        /// </summary>
        /// <param name="doc">The document containing unused regions.</param>
        /// <param name="handler">The handler which implements the IFieldMergingCallback interface
        /// and defines the logic to be applied to each unmerged region.</param>
        public void ExecuteCustomLogicOnEmptyRegions(Document doc, IFieldMergingCallback handler)
        {
            // Pass null to handle all regions found in the document.
            ExecuteCustomLogicOnEmptyRegions(doc, handler, null); 
        }

        /// <summary>
        /// Applies logic defined in the passed handler class to specific unused regions in the document as defined in regionsList.
        /// This allows controlling how unused regions are handled in the document manually.
        /// </summary>
        /// <param name="doc">The document containing unused regions.</param>
        /// <param name="handler">The handler which implements the IFieldMergingCallback interface and defines
        /// the logic to be applied to each unmerged region.</param>
        /// <param name="regionsList">A list of strings corresponding to the region names that are to be handled
        /// by the supplied handler class. Other regions encountered will not be handled and are removed automatically.</param>
        public void ExecuteCustomLogicOnEmptyRegions(Document doc, IFieldMergingCallback handler,
            List<string> regionsList)
        {
            // Certain regions can be skipped from applying logic to by not adding
            // the table name inside the CreateEmptyDataSource method. Enable this cleanup option, so any regions
            // which are not handled by the user's logic are removed automatically.
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedRegions;

            // Set the user's handler, which is called for each unmerged region.
            doc.MailMerge.FieldMergingCallback = handler;

            // Execute mail merge using the dummy dataset. The dummy data source contains each unmerged region's table names
            // in the document (excluding ones that the user may have specified to be skipped).
            // This will allow the handler to be called for each field in the unmerged regions.
            doc.MailMerge.ExecuteWithRegions(CreateDataSourceFromDocumentRegions(doc, regionsList));
        }
        //ExEnd:ExecuteCustomLogicOnEmptyRegions

        //ExStart:EmptyRegionsHandler 
        public class EmptyRegionsHandler : IFieldMergingCallback
        {
            /// <summary>
            /// Called for each field belonging to an unmerged region in the document.
            /// </summary>
            public void FieldMerging(FieldMergingArgs args)
            {
                // Change the text of each field of the ContactDetails region individually.
                if (args.TableName == "ContactDetails")
                {
                    // Set the text of the field based off the field name.
                    if (args.FieldName == "Name")
                        args.Text = "(No details found)";
                    else if (args.FieldName == "Number")
                        args.Text = "(N/A)";
                }

                // Remove the entire table of the Suppliers region. Also, check if the previous paragraph
                // before the table is a heading paragraph and remove that.
                if (args.TableName == "Suppliers")
                {
                    Table table = (Table) args.Field.Start.GetAncestor(NodeType.Table);

                    // Check if the table has been removed from the document already.
                    if (table.ParentNode != null)
                    {
                        // Try to find the paragraph which precedes the table before the table is removed from the document.
                        if (table.PreviousSibling != null && table.PreviousSibling.NodeType == NodeType.Paragraph)
                        {
                            Paragraph previousPara = (Paragraph) table.PreviousSibling;
                            if (IsHeadingParagraph(previousPara))
                                previousPara.Remove();
                        }

                        table.Remove();
                    }
                }
            }

            /// <summary>
            /// Returns true if the paragraph uses any Heading style, e.g., Heading 1 to Heading 9.
            /// </summary>
            private bool IsHeadingParagraph(Paragraph para)
            {
                return para.ParagraphFormat.StyleIdentifier >= StyleIdentifier.Heading1 &&
                       para.ParagraphFormat.StyleIdentifier <= StyleIdentifier.Heading9;
            }

            public void ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // Do nothing.
            }
        }
        //ExEnd:EmptyRegionsHandler 

        public class EmptyRegionsHandlerMergeTable : IFieldMergingCallback
        {
            /// <summary>
            /// Called for each field belonging to an unmerged region in the document.
            /// </summary>
            public void FieldMerging(FieldMergingArgs args)
            {
                //ExStart:RemoveExtraParagraphs
                // Store the parent paragraph of the current field for easy access.
                Paragraph parentParagraph = args.Field.Start.ParentParagraph;

                // Define the logic to be used when the ContactDetails region is encountered.
                // The region is removed and replaced with a single line of text stating that there are no records.
                if (args.TableName == "ContactDetails")
                {
                    // Called for the first field encountered in a region. This can be used to execute logic on the first field
                    // in the region without needing to hard code the field name. Often the base logic is applied to the first field and 
                    // different logic for other fields. The rest of the fields in the region will have a null FieldValue.
                    if ((string) args.FieldValue == "FirstField")
                    {
                        FindReplaceOptions options = new FindReplaceOptions();
                        // Remove the "Name:" tag from the start of the paragraph.
                        parentParagraph.Range.Replace("Name:", string.Empty, options);
                        // Set the text of the first field to display a message stating that there are no records.
                        args.Text = "No records to display";
                    }
                    else
                    {
                        // We have already inserted our message in the paragraph belonging to the first field.
                        // The other paragraphs in the region will remain, so we want to remove these.
                        // A check is added to ensure that the paragraph has not been removed,
                        // which may happen if more than one field is included in a paragraph.
                        if (parentParagraph.ParentNode != null)
                            parentParagraph.Remove();
                    }
                }
                //ExEnd:RemoveExtraParagraphs

                //ExStart:MergeAllCells
                // Replace the unused region in the table with a "no records" message and merge all cells into one.
                if (args.TableName == "Suppliers")
                {
                    if ((string) args.FieldValue == "FirstField")
                    {
                        // We will use the first paragraph to display our message. Make it centered within the table.
                        // The other fields in other cells within the table will be merged and won't be displayed,
                        // so we don't need to do anything else with them.
                        parentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                        args.Text = "No records to display";
                    }

                    // Merge the cells of the table.
                    Cell cell = (Cell) parentParagraph.GetAncestor(NodeType.Cell);
                    if (cell != null)
                    {
                        cell.CellFormat.HorizontalMerge = cell.IsFirstCell ? CellMerge.First : CellMerge.Previous;
                    }
                }
                //ExEnd:MergeAllCells
            }

            public void ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // Do Nothing
            }
        }

        /// <summary>
        /// Returns the data used to merge the document.
        /// This dataset purposely contains only rows for the StoreDetails region and only a select few for the child region. 
        /// </summary>
        private DataSet GetDataSource()
        {
            DataSet data = new DataSet();
            DataTable storeDetails = new DataTable("StoreDetails");
            DataTable contactDetails = new DataTable("ContactDetails");

            contactDetails.Columns.Add("ID");
            contactDetails.Columns.Add("Name");
            contactDetails.Columns.Add("Number");

            storeDetails.Columns.Add("ID");
            storeDetails.Columns.Add("Name");
            storeDetails.Columns.Add("Address");
            storeDetails.Columns.Add("City");
            storeDetails.Columns.Add("Country");

            storeDetails.Rows.Add("0", "Hungry Coyote Import Store", "2732 Baker Blvd", "Eugene", "USA");
            storeDetails.Rows.Add("1", "Great Lakes Food Market", "City Center Plaza, 516 Main St.", "San Francisco",
                "USA");

            contactDetails.Rows.Add("0", "Thomas Hardy", "(206) 555-9857 ext 237");
            contactDetails.Rows.Add("0", "Elizabeth Brown", "(206) 555-9857 ext 764");

            data.Tables.Add(storeDetails);
            data.Tables.Add(contactDetails);

            data.Relations.Add(storeDetails.Columns["ID"], contactDetails.Columns["ID"]);

            return data;
        }

        private readonly DataTable orderTable = null;
        private readonly DataTable itemTable = null;

        private void DisableForeignKeyConstraints(DataSet dataSet)
        {
            //ExStart:DisableForeignKeyConstraints
            dataSet.Relations.Add(new DataRelation("OrderToItem", orderTable.Columns["Order_Id"],
                itemTable.Columns["Order_Id"], false));
            //ExEnd:DisableForeignKeyConstraints
        }

        private void CreateDataRelation(DataSet dataSet)
        {
            //ExStart:CreateDataRelation
            dataSet.Relations.Add(new DataRelation("OrderToItem", orderTable.Columns["Order_Id"],
                itemTable.Columns["Order_Id"]));
            //ExEnd:CreateDataRelation
        }
    }
}