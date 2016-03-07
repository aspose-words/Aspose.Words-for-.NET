// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
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
