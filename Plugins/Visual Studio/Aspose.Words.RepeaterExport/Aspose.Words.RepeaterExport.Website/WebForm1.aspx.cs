using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Aspose.Words.RepeaterExport.Website
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                BindRepeater();
            }
        }
        private void BindRepeater()
        {
            ArrayList values = new ArrayList();

            values.Add("Apple");
            values.Add("Orange");
            values.Add("Pear");
            values.Add("Banana");
            values.Add("Grape");

            // Set the DataSource of the Repeater. 
            ExportRepeaterToWord1.ExportDataSource = values;
            ExportRepeaterToWord1.DataSource = values;
            ExportRepeaterToWord1.DataBind();
        }

        
    }
}