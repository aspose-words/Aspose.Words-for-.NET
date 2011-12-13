//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Data;
using System.Web.UI.WebControls;
using Aspose.Words;

namespace AjaxGenerateDoc
{
    /// <summary>
    /// Shows how to invoke Aspose.Words for generating a document with data from a GridView control. 
    /// In this example full post back is used.
    /// </summary>
    public partial class ExampleUsingPostBackTrigger : System.Web.UI.Page
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
        /// Generate file when "generate" row command occurs.
        /// </summary>
        protected void GridView1_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            if (e.CommandName == "generate")
            {
                int index = Convert.ToInt32(e.CommandArgument);
                GridViewRow row = GridView1.Rows[index];
                string name = row.Cells[1].Text;
                string company = row.Cells[2].Text;

                //Create a new document.
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);

                // Fill the document with custom data.
                builder.Writeln(String.Format("Hello {0} from {1}!", name, company));

                //Send created document to a client browser.
                doc.Save(Response, "out.doc", ContentDisposition.Attachment, null);

                Response.End();
            }
        }
    }    
}
