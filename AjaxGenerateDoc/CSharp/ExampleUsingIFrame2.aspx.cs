//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Data;
using System.Web.UI.WebControls;

namespace AjaxGenerateDoc
{
    /// <summary>
    /// Shows how to invoke Aspose.Words for generating a document with data from a GridView control. 
    /// In this example IFrame is used.
    /// </summary>
    public partial class ExampleUsingIFrame2 : System.Web.UI.Page
    {
        /// <summary>
        /// Fill GridView with data.
        /// </summary>
        protected void Page_Load(object sender, EventArgs e)
        {
            DataTable table = new DataTable();
            table.Columns.Add("Name");
            table.Columns.Add("Company");

            DataRow row1 = table.NewRow();
            row1["Name"] = "Alexey";
            row1["Company"] = "Aspose";
            table.Rows.Add(row1);

            DataRow row2 = table.NewRow();
            row2["Name"] = "Ravi";
            row2["Company"] = "Yolocounty";
            table.Rows.Add(row2);

            GridView1.DataSource = table;
            GridView1.DataBind();

        }

        /// <summary>
        /// Adds an onClick script that will invoke document generation (in an IFrame).
        /// </summary>
        protected void GridView1_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            // We make IFrame to invoke GenerateFile.aspx and pass parameters using a query string.
            string script =
                "var iframe = document.createElement('iframe'); " +
                "iframe.src = 'GenerateFile.aspx?name={0}&company={1}'; " +
                "iframe.style.display = 'none'; " +
                "document.body.appendChild(iframe);";

            script = String.Format(script, e.Row.Cells[1].Text, e.Row.Cells[2].Text);

            e.Row.Cells[0].Attributes.Add("onclick", script);
        }
    }    
}
