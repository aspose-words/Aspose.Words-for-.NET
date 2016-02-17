// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
/// <summary>
/// Represents a facade object for a column of a table in a Microsoft Word document.
/// </summary>
internal class Column
{
    private Column(Table table, int columnIndex)
    {
        if (table == null)
            throw new ArgumentException("table");

        mTable = table;
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
        get
        {
            return (Cell[])GetColumnCells().ToArray(typeof(Cell));
        }
    }

    /// <summary>
    /// Returns the index of the given cell in the column.
    /// </summary>
    public int IndexOf(Cell cell)
    {
        return GetColumnCells().IndexOf(cell);
    }

    /// <summary>
    /// Inserts a brand new column before this column into the table.
    /// </summary>
    public Column InsertColumnBefore()
    {
        Cell[] columnCells = Cells;

        if (columnCells.Length == 0)
            throw new ArgumentException("Column must not be empty");

        // Create a clone of this column.
        foreach (Cell cell in columnCells)
            cell.ParentRow.InsertBefore(cell.Clone(false), cell);

        // This is the new column.
        Column column = new Column(columnCells[0].ParentRow.ParentTable, mColumnIndex);

        // We want to make sure that the cells are all valid to work with (have at least one paragraph).
        foreach (Cell cell in column.Cells)
            cell.EnsureMinimum();

        // Increase the index which this column represents since there is now one extra column infront.
        mColumnIndex++;

        return column;
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

        foreach (Row row in mTable.Rows)
        {
            Cell cell = row.Cells[mColumnIndex];
            if (cell != null)
                columnCells.Add(cell);
        }

        return columnCells;
    }

    private int mColumnIndex;
    private Table mTable;
}
