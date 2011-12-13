//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using Aspose.Words;

namespace AjaxGenerateDoc
{
    /// <summary>
    /// This page is called inside an IFrame to generate a Microsoft Word document.
    /// 
    /// If the caller passes two parameters on the query string "name" and "company",
    /// they will be inserted into the generated document.
    /// </summary>
    public partial class GenerateFile : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string name = string.Empty;
            string company = string.Empty;

            if (Request.Params["name"] != null)
                name = Request.Params["name"];

            if (Request.Params["company"] != null)
                company = Request.Params["company"];

            //Create a new document.
            Document doc = new Document();

            //Fill the document with custom data.
            DocumentBuilder builder = new DocumentBuilder(doc);
            if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(company))
            {
                builder.Writeln("Hello World!");
            }
            else
            {
                builder.Writeln(String.Format("Hello {0} from {1}!", name, company));
            }

            //This delay is just for a demo! To simulate a delay when building a very complex document.
            System.Threading.Thread.Sleep(2000);

            // Let the caller know we have finished.
            Session["Completed"] = true;

            //Send the document to the browser.
            doc.Save(Response, "out.doc", ContentDisposition.Attachment, null);

            Response.End();
        }
    }
}
