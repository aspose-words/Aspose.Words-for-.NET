using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Collections;
using System.Data;

namespace Aspose.Words.RepeaterExport.Website
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                BindGrid();
            }
        }
        private void BindGrid()
        {
            DataTable dataTable = new DataTable("Products");
            dataTable.Columns.Add("Product ID", typeof(Int32));
            dataTable.Columns.Add("Product Name", typeof(string));
            dataTable.Columns.Add("Units In Stock", typeof(Int32));


            for (int index = 0; index < 10; index++)
            {
                DataRow dr = dataTable.NewRow();
                dr[0] = index;
                dr[1] = dr[0] + " - Name";
                dr[2] = index + 5;
                dataTable.Rows.Add(dr);
            }

            // Set the DataSource of the Repeater. 
            ExportRepeaterToWord1.ExportDataSource = dataTable;
            ExportRepeaterToWord1.DataSource = dataTable;
            ExportRepeaterToWord1.DataBind();
        }
    }
}