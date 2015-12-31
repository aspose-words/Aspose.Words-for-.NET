using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
namespace Aspose.Words.ListViewExport.Website
{
    public partial class Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                BindListView();
            }
        }

        private void BindListView()
        {
            DataTable dataTable = new DataTable("Products");
            dataTable.Columns.Add("Product ID", typeof(Int32));
            dataTable.Columns.Add("Product Name", typeof(string));
            dataTable.Columns.Add("Units In Stock", typeof(Int32));


            for (int index = 1; index <= 50; index++)
            {
                DataRow dr = dataTable.NewRow();
                dr[0] = index;
                dr[1] = "Product - " + dr[0];
                dr[2] = index + 5;
                dataTable.Rows.Add(dr);
            }


            ExportListViewToWord1.ExportDataSource = dataTable;
            ExportListViewToWord1.DataSource = dataTable;
            ExportListViewToWord1.DataBind();
        }

        protected void ExportListViewToWord1_PagePropertiesChanging(object sender, PagePropertiesChangingEventArgs e)
        {
            DataPager objDataPager1 = (DataPager)ExportListViewToWord1.FindControl("DataPager1");
            if (objDataPager1 != null)
            {
                objDataPager1.SetPageProperties(e.StartRowIndex, e.MaximumRows, false);
                BindListView();
            }
        }


        public override void VerifyRenderingInServerForm(Control control)
        { }

    }
}