
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Diagnostics;
using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Tables
{
    class MergedCells
    {
        public static void Run()
        {            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithTables();
            CheckCellsMerged(dataDir);
            // The below method shows how to create a table with two rows with cells in the first row horizontally merged.
            HorizontalMerge(dataDir);
            // The below method shows how to create a table with two columns with cells merged vertically in the first column.
            VerticalMerge(dataDir);
            // The below method shows how to merges the range of cells between the two specified cells.   
            MergeCellRange(dataDir);
            // Show how to prints the horizontal and vertical merge of a cell.
            PrintHorizontalAndVerticalMerged(dataDir);
        }
        public static void CheckCellsMerged(string dataDir)
        {
            //ExStart:CheckCellsMerged 
            Document doc = new Document(dataDir + "Table.MergedCells.doc");

            // Retrieve the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    Console.WriteLine(PrintCellMergeType(cell));
                }
            }
            //ExEnd:CheckCellsMerged 
        }
        //ExStart:PrintCellMergeType 
        public static string PrintCellMergeType(Cell cell)
        {
            bool isHorizontallyMerged = cell.CellFormat.HorizontalMerge != CellMerge.None;
            bool isVerticallyMerged = cell.CellFormat.VerticalMerge != CellMerge.None;
            string cellLocation = string.Format("R{0}, C{1}", cell.ParentRow.ParentTable.IndexOf(cell.ParentRow) + 1, cell.ParentRow.IndexOf(cell) + 1);

            if (isHorizontallyMerged && isVerticallyMerged)
                return string.Format("The cell at {0} is both horizontally and vertically merged", cellLocation);
            else if (isHorizontallyMerged)
                return string.Format("The cell at {0} is horizontally merged.", cellLocation);
            else if (isVerticallyMerged)
                return string.Format("The cell at {0} is vertically merged", cellLocation);
            else
                return string.Format("The cell at {0} is not merged", cellLocation);
        }
        //ExEnd:PrintCellMergeType
        public static void VerticalMerge( string dataDir)
        {
            //ExStart:VerticalMerge           
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.First;
            builder.Write("Text in merged cells.");

            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Text in one cell");
            builder.EndRow();

            builder.InsertCell();
            // This cell is vertically merged to the cell above and should be empty.
            builder.CellFormat.VerticalMerge = CellMerge.Previous;

            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Text in another cell");
            builder.EndRow();
            builder.EndTable();
            dataDir = dataDir + "Table.VerticalMerge_out_.doc";

            // Save the document to disk.
            doc.Save(dataDir);
            //ExEnd:VerticalMerge
            Console.WriteLine("\nTable created successfully with two columns with cells merged vertically in the first column.\nFile saved at " + dataDir);
        }
        public static void HorizontalMerge(string dataDir)
        {
            //ExStart:HorizontalMerge         
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Text in merged cells.");

            builder.InsertCell();
            // This cell is merged to the previous and should be empty.
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.EndRow();

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.Write("Text in one cell.");

            builder.InsertCell();
            builder.Write("Text in another cell.");
            builder.EndRow();
            builder.EndTable();
            dataDir = dataDir + "Table.HorizontalMerge_out_.doc";

            // Save the document to disk.
            doc.Save(dataDir);
            //ExEnd:HorizontalMerge
            Console.WriteLine("\nTable created successfully with cells in the first row horizontally merged.\nFile saved at " + dataDir);
            
        }
        public static void MergeCellRange(string dataDir)
        {
            //ExStart:MergeCellRange
            // Open the document
            Document doc = new Document(dataDir + "Table.Document.doc");

            // Retrieve the first table in the body of the first section.
            Table table = doc.FirstSection.Body.Tables[0];
           
            // We want to merge the range of cells found inbetween these two cells.
            Cell cellStartRange = table.Rows[2].Cells[2];
            Cell cellEndRange = table.Rows[3].Cells[3];

            // Merge all the cells between the two specified cells into one.
            MergeCells(cellStartRange, cellEndRange);            
            dataDir = dataDir + "Table.MergeCellRange_out_.doc";
            // Save the document.
            doc.Save(dataDir);
            //ExEnd:MergeCellRange
            Console.WriteLine("\nCells merged successfully.\nFile saved at " + dataDir);
            
        }
        public static void PrintHorizontalAndVerticalMerged(string dataDir)
        {
            //ExStart:PrintHorizontalAndVerticalMerged
            Document doc = new Document(dataDir + "Table.MergedCells.doc");

            //Create visitor
            SpanVisitor visitor = new SpanVisitor(doc);

            //Accept visitor
            doc.Accept(visitor);
            //ExEnd:PrintHorizontalAndVerticalMerged
            Console.WriteLine("\nHorizontal and vertical merged of a cell prints successfully.");
           
        }
        //ExStart:MergeCells
        internal static void MergeCells(Cell startCell, Cell endCell)
        {
            Table parentTable = startCell.ParentRow.ParentTable;

            // Find the row and cell indices for the start and end cell.
            Point startCellPos = new Point(startCell.ParentRow.IndexOf(startCell), parentTable.IndexOf(startCell.ParentRow));
            Point endCellPos = new Point(endCell.ParentRow.IndexOf(endCell), parentTable.IndexOf(endCell.ParentRow));
            // Create the range of cells to be merged based off these indices. Inverse each index if the end cell if before the start cell. 
            Rectangle mergeRange = new Rectangle( System.Math.Min(startCellPos.X, endCellPos.X), System.Math.Min(startCellPos.Y, endCellPos.Y),
                System.Math.Abs(endCellPos.X - startCellPos.X) + 1, System.Math.Abs(endCellPos.Y - startCellPos.Y) + 1);

            foreach (Row row in parentTable.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    Point currentPos = new Point(row.IndexOf(cell), parentTable.IndexOf(row));

                    // Check if the current cell is inside our merge range then merge it.
                    if (mergeRange.Contains(currentPos))
                    {
                        if (currentPos.X == mergeRange.X)
                            cell.CellFormat.HorizontalMerge = CellMerge.First;
                        else
                            cell.CellFormat.HorizontalMerge = CellMerge.Previous;

                        if (currentPos.Y == mergeRange.Y)
                            cell.CellFormat.VerticalMerge = CellMerge.First;
                        else
                            cell.CellFormat.VerticalMerge = CellMerge.Previous;
                    }
                }
            }
        }
        //ExEnd:MergeCells
        //ExStart:HorizontalAndVerticalMergeHelperClasses
        /// <summary>
        /// Helper class that contains collection of rowinfo for each row
        /// </summary>
        public class TableInfo
        {
            public List<RowInfo> Rows
            {
                get { return mRows; }
            }

            private List<RowInfo> mRows = new List<RowInfo>();
        }

        /// <summary>
        /// Helper class that contains collection of cellinfo for each cell
        /// </summary>
        public class RowInfo
        {
            public List<CellInfo> Cells
            {
                get { return mCells; }
            }

            private List<CellInfo> mCells = new List<CellInfo>();
        }

        /// <summary>
        /// Helper class that contains info about cell. currently here is only colspan and rowspan
        /// </summary>
        public class CellInfo
        {
            public CellInfo(int colSpan, int rowSpan)
            {
                mColSpan = colSpan;
                mRowSpan = rowSpan;
            }

            public int ColSpan
            {
                get { return mColSpan; }
            }

            public int RowSpan
            {
                get { return mRowSpan; }
            }

            private int mColSpan = 0;
            private int mRowSpan = 0;
        }

        public class SpanVisitor : DocumentVisitor
        {

            /// <summary>
            /// Creates new SpanVisitor instance
            /// </summary>
            /// <param name="doc">Is document which we should parse</param>
            public SpanVisitor(Document doc)
            {
                // Get collection of tables from the document
                mWordTables = doc.GetChildNodes(NodeType.Table, true);

                // Convert document to HTML
                // We will parse HTML to determine rowspan and colspan of each cell
                MemoryStream htmlStream = new MemoryStream();

                HtmlSaveOptions options = new HtmlSaveOptions();
                options.ImagesFolder = Path.GetTempPath();

                doc.Save(htmlStream, options);

                // Load HTML into the XML document
                XmlDocument xmlDoc = new XmlDocument();
                htmlStream.Position = 0;
                xmlDoc.Load(htmlStream);

                // Get collection of tables in the HTML document
                XmlNodeList tables = xmlDoc.DocumentElement.SelectNodes("//table");

                foreach (XmlNode table in tables)
                {
                    TableInfo tableInf = new TableInfo();
                    // Get collection of rows in the table
                    XmlNodeList rows = table.SelectNodes("tr");

                    foreach (XmlNode row in rows)
                    {
                        RowInfo rowInf = new RowInfo();

                        // Get collection of cells
                        XmlNodeList cells = row.SelectNodes("td");

                        foreach (XmlNode cell in cells)
                        {
                            // Determine row span and colspan of the current cell
                            XmlAttribute colSpanAttr = cell.Attributes["colspan"];
                            XmlAttribute rowSpanAttr = cell.Attributes["rowspan"];

                            int colSpan = colSpanAttr == null ? 0 : Int32.Parse(colSpanAttr.Value);
                            int rowSpan = rowSpanAttr == null ? 0 : Int32.Parse(rowSpanAttr.Value);

                            CellInfo cellInf = new CellInfo(colSpan, rowSpan);
                            rowInf.Cells.Add(cellInf);
                        }

                        tableInf.Rows.Add(rowInf);
                    }

                    mTables.Add(tableInf);
                }
            }

            public override VisitorAction VisitCellStart(Tables.Cell cell)
            {
                // Determone index of current table
                int tabIdx = mWordTables.IndexOf(cell.ParentRow.ParentTable);

                // Determine index of current row
                int rowIdx = cell.ParentRow.ParentTable.IndexOf(cell.ParentRow);

                // And determine index of current cell
                int cellIdx = cell.ParentRow.IndexOf(cell);

                // Determine colspan and rowspan of current cell
                int colSpan = 0;
                int rowSpan = 0;
                if (tabIdx < mTables.Count &&
                    rowIdx < mTables[tabIdx].Rows.Count &&
                    cellIdx < mTables[tabIdx].Rows[rowIdx].Cells.Count)
                {
                    colSpan = mTables[tabIdx].Rows[rowIdx].Cells[cellIdx].ColSpan;
                    rowSpan = mTables[tabIdx].Rows[rowIdx].Cells[cellIdx].RowSpan;
                }

                Console.WriteLine("{0}.{1}.{2} colspan={3}\t rowspan={4}", tabIdx, rowIdx, cellIdx, colSpan, rowSpan);

                return VisitorAction.Continue;
            }
            private List<TableInfo> mTables = new List<TableInfo>();
            private NodeCollection mWordTables = null;
        }
        //ExEnd:HorizontalAndVerticalMergeHelperClasses
 
    }
}
