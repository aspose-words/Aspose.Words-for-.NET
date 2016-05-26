// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections;
using System.Text;

using Aspose.Words;
using Aspose.Words.Tables;

using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExTableColumn : ApiExampleBase
    {
        //ExStart
        //ExId:ColumnFacade
        //ExSummary:Demonstrates a facade object for working with a column of a table.
        /// <summary>
        /// Represents a facade object for a column of a table in a Microsoft Word document.
        /// </summary>
        public class Column
        {
            private Column(Table table, int columnIndex)
            {
                if (table == null)
                    throw new ArgumentException("table");

                this.mTable = table;
                this.mColumnIndex = columnIndex;
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
                get
                {
                    return (Cell[])this.GetColumnCells().ToArray(typeof(Cell));
                }
            }

            /// <summary>
            /// Returns the index of the given cell in the column.
            /// </summary>
            public int IndexOf(Cell cell)
            {
                return this.GetColumnCells().IndexOf(cell);
            }

            /// <summary>
            /// Inserts a brand new column before this column into the table.
            /// </summary>
            public Column InsertColumnBefore()
            {
                Cell[] columnCells = this.Cells;

                if (columnCells.Length == 0)
                    throw new ArgumentException("Column must not be empty");

                // Create a clone of this column.
                foreach (Cell cell in columnCells)
                    cell.ParentRow.InsertBefore(cell.Clone(false), cell);

                // This is the new column.
                Column column = new Column(columnCells[0].ParentRow.ParentTable, this.mColumnIndex);

                // We want to make sure that the cells are all valid to work with (have at least one paragraph).
                foreach (Cell cell in column.Cells)
                    cell.EnsureMinimum();

                // Increase the index which this column represents since there is now one extra column infront.
                this.mColumnIndex++;

                return column;
            }

            /// <summary>
            /// Removes the column from the table.
            /// </summary>
            public void Remove()
            {
                foreach (Cell cell in this.Cells)
                    cell.Remove();
            }

            /// <summary>
            /// Returns the text of the column. 
            /// </summary>
            public string ToTxt()
            {
                StringBuilder builder = new StringBuilder();

                foreach (Cell cell in this.Cells)
                    builder.Append(cell.ToString(SaveFormat.Text));

                return builder.ToString();
            }

            /// <summary>
            /// Provides an up-to-date collection of cells which make up the column represented by this facade.
            /// </summary>
            private ArrayList GetColumnCells()
            {
                ArrayList columnCells = new ArrayList();

                foreach (Row row in this.mTable.Rows)
                {
                    Cell cell = row.Cells[this.mColumnIndex];
                    if (cell != null)
                        columnCells.Add(cell);
                }

                return columnCells;
            }

            private int mColumnIndex;
            private Table mTable;
        }
        //ExEnd

        [Test]
        public void RemoveColumnFromTable()
        {
            //ExStart
            //ExId:RemoveTableColumn
            //ExSummary:Shows how to remove a column from a table in a document.
            Document doc = new Document(MyDir + "Table.Document.doc");
            Table table = (Table)doc.GetChild(NodeType.Table, 1, true);

            // Get the third column from the table and remove it.
            Column column = Column.FromIndex(table, 2);
            column.Remove();
            //ExEnd

            doc.Save(MyDir + @"\Artifacts\Table.RemoveColumn.doc");

            Assert.AreEqual(16, table.GetChildNodes(NodeType.Cell, true).Count);
            Assert.AreEqual("Cell 3 contents", table.Rows[2].Cells[2].ToString(SaveFormat.Text).Trim());
            Assert.AreEqual("Cell 3 contents", table.LastRow.Cells[2].ToString(SaveFormat.Text).Trim());
        }

        [Test]
        public void InsertNewColumnIntoTable()
        {
            Document doc = new Document(MyDir + "Table.Document.doc");
            Table table = (Table)doc.GetChild(NodeType.Table, 1, true);

            //ExStart
            //ExId:InsertNewColumn
            //ExSummary:Shows how to insert a blank column into a table.
            // Get the second column in the table.
            Column column = Column.FromIndex(table, 1);

            // Create a new column to the left of this column.
            // This is the same as using the "Insert Column Before" command in Microsoft Word.
            Column newColumn = column.InsertColumnBefore();

            // Add some text to each of the column cells.
            foreach (Cell cell in newColumn.Cells)
                cell.FirstParagraph.AppendChild(new Run(doc, "Column Text " + newColumn.IndexOf(cell)));
            //ExEnd

            doc.Save(MyDir + @"\Artifacts\Table.InsertColumn.doc");

            Assert.AreEqual(24, table.GetChildNodes(NodeType.Cell, true).Count);
            Assert.AreEqual("Column Text 0", table.FirstRow.Cells[1].ToString(SaveFormat.Text).Trim());
            Assert.AreEqual("Column Text 3", table.LastRow.Cells[1].ToString(SaveFormat.Text).Trim());
        }

        [Test]
        public void TableColumnToTxt()
        {
            Document doc = new Document(MyDir + "Table.Document.doc");
            Table table = (Table)doc.GetChild(NodeType.Table, 1, true);

            //ExStart
            //ExId:TableColumnToTxt
            //ExSummary:Shows how to get the plain text of a table column.
            // Get the first column in the table.
            Column column = Column.FromIndex(table, 0);

            // Print the plain text of the column to the screen.
            Console.WriteLine(column.ToTxt());
            //ExEnd

            Assert.AreEqual("\r\nRow 1\r\nRow 2\r\nRow 3\r\n", column.ToTxt());
        }
    }
}
