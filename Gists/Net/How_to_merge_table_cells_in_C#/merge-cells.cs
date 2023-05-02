// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
internal void MergeCells(Cell startCell, Cell endCell)
{
    Table parentTable = startCell.ParentRow.ParentTable;

    // Find the row and cell indices for the start and end cell.
    Point startCellPos = new Point(startCell.ParentRow.IndexOf(startCell),
        parentTable.IndexOf(startCell.ParentRow));
    Point endCellPos = new Point(endCell.ParentRow.IndexOf(endCell), parentTable.IndexOf(endCell.ParentRow));

    // Create a range of cells to be merged based on these indices.
    // Inverse each index if the end cell is before the start cell.
    Rectangle mergeRange = new Rectangle(Math.Min(startCellPos.X, endCellPos.X),
        Math.Min(startCellPos.Y, endCellPos.Y),
        Math.Abs(endCellPos.X - startCellPos.X) + 1, Math.Abs(endCellPos.Y - startCellPos.Y) + 1);

    foreach (Row row in parentTable.Rows)
    {
        foreach (Cell cell in row.Cells)
        {
            Point currentPos = new Point(row.IndexOf(cell), parentTable.IndexOf(row));

            // Check if the current cell is inside our merge range, then merge it.
            if (mergeRange.Contains(currentPos))
            {
                cell.CellFormat.HorizontalMerge = currentPos.X == mergeRange.X ? CellMerge.First : CellMerge.Previous;

                cell.CellFormat.VerticalMerge = currentPos.Y == mergeRange.Y ? CellMerge.First : CellMerge.Previous;
            }
        }
    }
}
