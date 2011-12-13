//ExStart
//ExId:Azure_ConvertDocumentSimple
//ExSummary:Shows how to convert a document in Windows Azure.
using Aspose.Words;
using Aspose.Words.Saving;
using System.Web;
using System;
using System.IO;

namespace WebRole
{
    /// <summary>
    /// This demo shows how to use Aspose.Words for .NET inside a WebRole in a simple
    /// Windows Azure application. There is just one ASP.NET page that provides a user
    /// interface to convert a document from one format to another.
    /// </summary>
    public partial class _Default : System.Web.UI.Page
    {
        protected void SubmitButton_Click(object sender, EventArgs e)
        {
            HttpPostedFile postedFile = SrcFileUpload.PostedFile;

            if (postedFile.ContentLength == 0)
                throw new Exception("There was no document uploaded.");

            if (postedFile.ContentLength > 512 * 1024)
                throw new Exception("The uploaded document is too big. This demo limits the file size to 512Kb.");

            // Create a suitable file name for the converted document.
            string dstExtension = DstFormatDropDownList.SelectedValue;
            string dstFileName = Path.GetFileName(postedFile.FileName) + "_Converted." + dstExtension;
            SaveFormat dstFormat = FileFormatUtil.ExtensionToSaveFormat(dstExtension);

            // Convert the document and send to the browser.
            Document doc = new Document(postedFile.InputStream);
            doc.Save(Response, dstFileName, ContentDisposition.Inline, SaveOptions.CreateSaveOptions(dstFormat));
            // Required. Otherwise DOCX cannot be opened on the client (probably not all data sent
            // or some extra data sent in the response).
            Response.End();
        }

        static _Default()
        {
            // Uncomment this code and embed your license file as a resource in this project and this code 
            // will find and activate the license. Aspose.Wods licensing needs to execute only once
            // before any Document instance is created and a static ctor is a good place.
            //
            // Aspose.Words.License l = new Aspose.Words.License();
            // l.SetLicense("Aspose.Words.lic");
        }
    }
}
//ExEnd