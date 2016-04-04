// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
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
