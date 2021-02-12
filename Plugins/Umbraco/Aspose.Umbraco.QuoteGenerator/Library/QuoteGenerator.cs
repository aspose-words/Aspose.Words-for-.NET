using System;
using System.Linq;
using System.Web.UI.WebControls;
using System.Data;
using Aspose.Words;

namespace Aspose.UmbracoQuoteGenerator
{
    public class QuoteGenerator
    {
        public static void Run()
        {
            try
            {

            }
            catch (Exception exc)
            {
                throw exc;
            }
        }

        // Populate the VAT dropdown list.
        public static void PopulateVATDropdownList(ref DropDownList ddlVAT, System.Web.SessionState.HttpSessionState currentSession)
        {
            try
            {
                // If session already contains the list items, then we do not need to execute loop and items.
                if (currentSession["VATListItems"] != null)
                {
                    // Extract items from session.
                    ddlVAT.Items.AddRange(((ListItemCollection)currentSession["VATListItems"]).Cast<System.Web.UI.WebControls.ListItem>().ToArray());
                }
                else
                {
                    // Creating items for VAT %age from 1 to 20 with decimal places 1 to 9,
                    // in this way we will have items like (e.g 1%, 1.1%, 1.2%.........19.9%, 20%).
                    // Outer loop for 1 to 20 items.
                    for (int i = 0; i < 20; i++)
                    {
                        // Inner loop to create decimal items 1 to 9 for each outer loop value.
                        for (int j = 0; j < 10; j++)
                        {
                            // NOTE: (j == 0 ? "" : "." + j.ToString()) skip and allow to add start value.
                            ddlVAT.Items.Add(new ListItem(i.ToString() + (j == 0 ? "" : "." + j.ToString()) + "%", i.ToString() + (j == 0 ? "" : "." + j.ToString())));
                        }
                    }
                    // Adding last item as loops created max 19.9% item.
                    ddlVAT.Items.Add(new ListItem("20%", "20"));

                    // Adding list items to session, caching it to re-use.
                    currentSession["VATListItems"] = ddlVAT.Items;
                }
            }
            catch (Exception exc)
            {
                throw exc;
            }
        }

        // Populate the VAT dropdown list.
        public static DataSet GetDataSetForGridView(System.Web.SessionState.HttpSessionState currentSession)
        {
            try
            {
                // If session already contains the list items, then we do not need to execute loop and items.
                if (currentSession["ProductTable"] != null)
                {
                    // Extract items from session.
                    return (DataSet)currentSession["ProductTable"];
                }
                else
                {
                    // Create a new DataSet and DataTable objects to be used for mail merge.
                    DataSet data = new DataSet();
                    DataTable productTable = new DataTable("Products");

                    // Add columns for the productTable table.
                    productTable.Columns.Add("ProductId");
                    productTable.Columns.Add("ProductDescription");
                    productTable.Columns.Add("Price");
                    productTable.Columns.Add("Quantity");
                    productTable.Columns.Add("TotalBeforVAT");
                    productTable.Columns.Add("VATPercent");
                    productTable.Columns.Add("VATAmount");
                    productTable.Columns.Add("TotalAmount");

                    // Include the tables in the DataSet.
                    data.Tables.Add(productTable);

                    // Add dataset to session, caching it to re-use.
                    currentSession["ProductTable"] = data;

                    return data;
                }
            }
            catch (Exception exc)
            {
                throw exc;
            }
        }

        // Populate the VAT dropdown list.
        public static Document GetUnmergedTemplateObject(string templatePath, System.Web.SessionState.HttpSessionState currentSession)
        {
            try
            {
                Document doc = new Document(templatePath);
                return doc;
                // If session already contains the list items, then we do not need to execute loop and items.
                if (currentSession["TemplateObject"] != null)
                {
                    // Extract items from session.
                    doc = (Document)currentSession["TemplateObject"];

                    // If there are multiple templates and need to cache and use then must verify what session exactly have in.
                    if (templatePath.Contains(doc.OriginalFileName))
                    {
                        return doc;
                    }
                    else
                    {
                        // Create a new document object with un-merged template file, caching it to re-use.
                        doc = new Document(templatePath);

                        // Add template object to session.
                        currentSession["TemplateObject"] = doc;

                        return doc;
                    }
                }
                else
                {
                    // Create a new document object with un-merged template file, caching it to re-use.
                    doc = new Document(templatePath);

                    // Add template object to session.
                    currentSession["TemplateObject"] = doc;

                    return doc;
                }
            }
            catch (Exception exc)
            {
                throw exc;
            }
        }

        // Get file export types/extenssions.
        public static string GetSaveFormat(string format)
        {
            try
            {
                string saveOption = SaveFormat.Pdf.ToString();

                // There are many document formats supported, check the "SaveFormat" property for more.
                switch (format)
                {
                    case "Pdf":
                        saveOption = SaveFormat.Pdf.ToString(); break;
                    case "Doc":
                        saveOption = SaveFormat.Doc.ToString(); break;
                    case "Docx":
                        saveOption = SaveFormat.Docx.ToString(); break;
                    case "Odt":
                        saveOption = SaveFormat.Odt.ToString(); break;
                    case "Xps":
                        saveOption = SaveFormat.Xps.ToString(); break;
                    case "Tiff":
                        saveOption = SaveFormat.Tiff.ToString(); break;
                    case "Png":
                        saveOption = SaveFormat.Png.ToString(); break;
                    case "Jpeg":
                        saveOption = SaveFormat.Jpeg.ToString(); break;

                }

                return saveOption;
            }
            catch (Exception exc)
            {
                throw exc;
            }
        }
    }
}