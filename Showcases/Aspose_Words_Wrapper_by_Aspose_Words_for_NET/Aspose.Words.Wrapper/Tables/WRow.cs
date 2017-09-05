using Aspose.Words.Tables;

namespace Aspose.Words.Wrapper.Tables
{
    public class WRow : Row
    {
        public WRow(DocumentBase document) : base(document)
        {
        }

        public Cell GetCell(int cellNumber)
        {
            return cellNumber >= 0 && cellNumber < Cells.Count 
                ? Cells[cellNumber] 
                : null;
        }
    }
}
