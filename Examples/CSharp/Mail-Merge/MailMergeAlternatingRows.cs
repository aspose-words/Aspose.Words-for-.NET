using Aspose.Words;
//ExStart:MailMergingNamespace
using Aspose.Words.MailMerging;
//ExEnd:MailMergingNamespace
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Drawing;

namespace Aspose.Words.Examples.CSharp.Mail_Merge
{
    class MailMergeAlternatingRows
    {
        public static void Run()
        {
            //ExStart:MailMergeAlternatingRows           
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting(); 
            Document doc = new Document(dataDir + "MailMerge.AlternatingRows.doc");

            // Add a handler for the MergeField event.
            doc.MailMerge.FieldMergingCallback = new HandleMergeFieldAlternatingRows();

            // Execute mail merge with regions.
            DataTable dataTable = GetSuppliersDataTable();
            doc.MailMerge.ExecuteWithRegions(dataTable);
            dataDir = dataDir + "MailMerge.AlternatingRows_out_.doc";
            doc.Save(dataDir);
            //ExEnd:MailMergeAlternatingRows
            Console.WriteLine("\nMail merge alternative rows performed successfully.\nFile saved at " + dataDir);
        }
        //ExStart:HandleMergeFieldAlternatingRows
        private class HandleMergeFieldAlternatingRows : IFieldMergingCallback
        {
            /// <summary>
            /// Called for every merge field encountered in the document.
            /// We can either return some data to the mail merge engine or do something
            /// else with the document. In this case we modify cell formatting.
            /// </summary>
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
            {
                if (mBuilder == null)
                    mBuilder = new DocumentBuilder(e.Document);

                // This way we catch the beginning of a new row.
                if (e.FieldName.Equals("CompanyName"))
                {
                    // Select the color depending on whether the row number is even or odd.
                    Color rowColor;
                    if (IsOdd(mRowIdx))
                        rowColor = Color.FromArgb(213, 227, 235);
                    else
                        rowColor = Color.FromArgb(242, 242, 242);

                    // There is no way to set cell properties for the whole row at the moment,
                    // so we have to iterate over all cells in the row.
                    for (int colIdx = 0; colIdx < 4; colIdx++)
                    {
                        mBuilder.MoveToCell(0, mRowIdx, colIdx, 0);
                        mBuilder.CellFormat.Shading.BackgroundPatternColor = rowColor;
                    }

                    mRowIdx++;
                }
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // Do nothing.
            }

            private DocumentBuilder mBuilder;
            private int mRowIdx;
        }     
        /// <summary>
        /// Returns true if the value is odd; false if the value is even.
        /// </summary>
        private static bool IsOdd(int value)
        {
            // The code is a bit complex, but otherwise automatic conversion to VB does not work.
            return ((value / 2) * 2).Equals(value);
        }      
        /// <summary>
        /// Create DataTable and fill it with data.
        /// In real life this DataTable should be filled from a database.
        /// </summary>
        private static DataTable GetSuppliersDataTable()
        {
            DataTable dataTable = new DataTable("Suppliers");
            dataTable.Columns.Add("CompanyName");
            dataTable.Columns.Add("ContactName");
            for (int i = 0; i < 10; i++)
            {
                DataRow datarow = dataTable.NewRow();
                dataTable.Rows.Add(datarow);
                datarow[0] = "Company " + i.ToString();
                datarow[1] = "Contact " + i.ToString();
            }
            return dataTable;
        }
        //ExEnd:HandleMergeFieldAlternatingRows
    }
}
