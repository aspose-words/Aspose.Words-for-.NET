using Aspose.Words.Tables;

namespace Aspose.Words.Wrapper.Tables
{
    public class WTable : Table
    {
        public WTable(DocumentBase document) : base(document)
        {
        }

        public Row GetRow(int rowNumber)
        {
            return rowNumber >= 0 && rowNumber < Rows.Count 
                ? Rows[rowNumber]
                : null;
        }
    }
}
