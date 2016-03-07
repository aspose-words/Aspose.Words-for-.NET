' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
''' <summary>
''' Create DataTable and fill it with data.
''' In real life this DataTable should be filled from a database.
''' </summary>
Private Shared Function GetSuppliersDataTable() As DataTable
    Dim dataTable As New DataTable("Suppliers")
    dataTable.Columns.Add("CompanyName")
    dataTable.Columns.Add("ContactName")
    For i As Integer = 0 To 9
        Dim datarow As DataRow = dataTable.NewRow()
        dataTable.Rows.Add(datarow)
        datarow(0) = "Company " + i.ToString()
        datarow(1) = "Contact " + i.ToString()
    Next
    Return dataTable
End Function
