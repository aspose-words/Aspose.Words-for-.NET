// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExTableColumn : ApiExampleBase
    {
        /// <summary>
        /// Represents a facade object for a column of a table in a Microsoft Word document.
        /// </summary>
        public class Column
        {
            private Column(Table table, int columnIndex)
            {
                mTable = table ?? throw new ArgumentException("table");
                mColumnIndex = columnIndex;
            }

            /// <summary>
            /// Returns a new column facade from the table and supplied zero-based index.
            /// </summary>
            public static Column FromIndex(Table table, int columnIndex)
            {
                return new Column(table, columnIndex);
            }

            /// <summary>
            /// Returns the cells which make up the column.
            /// </summary>
            public Cell[] Cells
            {
                get { return (Cell[]) GetColumnCells().ToArray(typeof(Cell)); }
            }

            /// <summary>
            /// Returns the index of the given cell in the column.
            /// </summary>
            public int IndexOf(Cell cell)
            {
                return GetColumnCells().IndexOf(cell);
            }

            /// <summary>
            /// Inserts a new column before this column into the table.
            /// </summary>
            public Column InsertColumnBefore()
            {
                Cell[] columnCells = Cells;

                if (columnCells.Length == 0)
                    throw new ArgumentException("Column must not be empty");

                // Create a clone of this column
                foreach (Cell cell in columnCells)
                    cell.ParentRow.InsertBefore(cell.Clone(false), cell);
                
                Column newColumn = new Column(columnCells[0].ParentRow.ParentTable, mColumnIndex);

                // We want to make sure that the cells are all valid to work with (have at least one paragraph).
                foreach (Cell cell in newColumn.Cells)
                    cell.EnsureMinimum();

                // Increment the index of this column represents since there is a new column before it.
                mColumnIndex++;

                return newColumn;
            }

            /// <summary>
            /// Removes the column from the table.
            /// </summary>
            public void Remove()
            {
                foreach (Cell cell in Cells)
                    cell.Remove();
            }

            /// <summary>
            /// Returns the text of the column. 
            /// </summary>
            public string ToTxt()
            {
                StringBuilder builder = new StringBuilder();

                foreach (Cell cell in Cells)
                    builder.Append(cell.ToString(SaveFormat.Text));

                return builder.ToString();
            }

            /// <summary>
            /// Provides an up-to-date collection of cells which make up the column represented by this facade.
            /// </summary>
            private ArrayList GetColumnCells()
            {
                ArrayList columnCells = new ArrayList();

                foreach (Row row in mTable.Rows.OfType<Row>())
                {
                    Cell cell = row.Cells[mColumnIndex];
                    if (cell != null)
                        columnCells.Add(cell);
                }

                return columnCells;
            }

            private int mColumnIndex;
            private readonly Table mTable;
        }
        
        [Test]
        public void RemoveColumnFromTable()
        {
            Document doc = new Document(MyDir + "Tables.docx");
            Table table = (Table) doc.GetChild(NodeType.Table, 1, true);

            Column column = Column.FromIndex(table, 2);
            column.Remove();
            
            doc.Save(ArtifactsDir + "TableColumn.RemoveColumn.doc");

            Assert.AreEqual(16, table.GetChildNodes(NodeType.Cell, true).Count);
            Assert.AreEqual("Cell 7 contents", table.Rows[2].Cells[2].ToString(SaveFormat.Text).Trim());
            Assert.AreEqual("Cell 11 contents", table.LastRow.Cells[2].ToString(SaveFormat.Text).Trim());
        }

        [Test]
        public void Insert()
        {
            Document doc = new Document(MyDir + "Tables.docx");
            Table table = (Table) doc.GetChild(NodeType.Table, 1, true);

            Column column = Column.FromIndex(table, 1);

            // Create a new column to the left of this column.
            // This is the same as using the "Insert Column Before" command in Microsoft Word.
            Column newColumn = column.InsertColumnBefore();

            // Add some text to each cell in the column.
            foreach (Cell cell in newColumn.Cells)
                cell.FirstParagraph.AppendChild(new Run(doc, "Column Text " + newColumn.IndexOf(cell)));
            
            doc.Save(ArtifactsDir + "TableColumn.Insert.doc");

            Assert.AreEqual(24, table.GetChildNodes(NodeType.Cell, true).Count);
            Assert.AreEqual("Column Text 0", table.FirstRow.Cells[1].ToString(SaveFormat.Text).Trim());
            Assert.AreEqual("Column Text 3", table.LastRow.Cells[1].ToString(SaveFormat.Text).Trim());
        }

        [Test]
        public void TableColumnToTxt()
        {
            Document doc = new Document(MyDir + "Tables.docx");
            Table table = (Table) doc.GetChild(NodeType.Table, 1, true);

            Column column = Column.FromIndex(table, 0);
            Console.WriteLine(column.ToTxt());

            Assert.AreEqual("\r\nRow 1\r\nRow 2\r\nRow 3\r\n", column.ToTxt());
        }
    }
}