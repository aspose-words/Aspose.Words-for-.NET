// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
/// <summary>
/// Helper class that contains collection of rowinfo for each row.
/// </summary>
public class TableInfo
{
    public List<RowInfo> Rows { get; } = new List<RowInfo>();
}

/// <summary>
/// Helper class that contains collection of cellinfo for each cell.
/// </summary>
public class RowInfo
{
    public List<CellInfo> Cells { get; } = new List<CellInfo>();
}

/// <summary>
/// Helper class that contains info about cell. currently here is only colspan and rowspan.
/// </summary>
public class CellInfo
{
    public CellInfo(int colSpan, int rowSpan)
    {
        ColSpan = colSpan;
        RowSpan = rowSpan;
    }

    public int ColSpan { get; }
    public int RowSpan { get; }
}

public class SpanVisitor : DocumentVisitor
{
    /// <summary>
    /// Creates new SpanVisitor instance.
    /// </summary>
    /// <param name="doc">
    /// Is document which we should parse.
    /// </param>
    public SpanVisitor(Document doc)
    {
        mWordTables = doc.GetChildNodes(NodeType.Table, true);

        // We will parse HTML to determine the rowspan and colspan of each cell.
        MemoryStream htmlStream = new MemoryStream();

        Aspose.Words.Saving.HtmlSaveOptions options = new Aspose.Words.Saving.HtmlSaveOptions
        {
            ImagesFolder = Path.GetTempPath()
        };

        doc.Save(htmlStream, options);

        // Load HTML into the XML document.
        XmlDocument xmlDoc = new XmlDocument();
        htmlStream.Position = 0;
        xmlDoc.Load(htmlStream);

        // Get collection of tables in the HTML document.
        XmlNodeList tables = xmlDoc.DocumentElement.GetElementsByTagName("table");

        foreach (XmlNode table in tables)
        {
            TableInfo tableInf = new TableInfo();
            // Get collection of rows in the table.
            XmlNodeList rows = table.SelectNodes("tr");

            foreach (XmlNode row in rows)
            {
                RowInfo rowInf = new RowInfo();
                // Get collection of cells.
                XmlNodeList cells = row.SelectNodes("td");

                foreach (XmlNode cell in cells)
                {
                    // Determine row span and colspan of the current cell.
                    XmlAttribute colSpanAttr = cell.Attributes["colspan"];
                    XmlAttribute rowSpanAttr = cell.Attributes["rowspan"];

                    int colSpan = colSpanAttr == null ? 0 : int.Parse(colSpanAttr.Value);
                    int rowSpan = rowSpanAttr == null ? 0 : int.Parse(rowSpanAttr.Value);

                    CellInfo cellInf = new CellInfo(colSpan, rowSpan);
                    rowInf.Cells.Add(cellInf);
                }

                tableInf.Rows.Add(rowInf);
            }

            mTables.Add(tableInf);
        }
    }

    public override VisitorAction VisitCellStart(Cell cell)
    {
        int tabIdx = mWordTables.IndexOf(cell.ParentRow.ParentTable);
        int rowIdx = cell.ParentRow.ParentTable.IndexOf(cell.ParentRow);
        int cellIdx = cell.ParentRow.IndexOf(cell);

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

    private readonly List<TableInfo> mTables = new List<TableInfo>();
    private readonly NodeCollection mWordTables;
}
