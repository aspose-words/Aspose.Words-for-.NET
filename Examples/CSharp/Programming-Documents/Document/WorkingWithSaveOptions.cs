using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    public class WorkingWithSaveOptions
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();

            UpdateLastSavedTimeProperty(dataDir);
            SetMeasureUnitForODT(dataDir);
        }

        public static void UpdateLastSavedTimeProperty(String dataDir)
        {
            // ExStart:UpdateLastSavedTimeProperty
            Document doc = new Document(dataDir + "Document.doc");

            OoxmlSaveOptions options = new OoxmlSaveOptions();
            options.UpdateLastSavedTimeProperty = true;

            dataDir = dataDir + "UpdateLastSavedTimeProperty_out.docx";

            // Save the document to disk.
            doc.Save(dataDir, options);
            // ExEnd:UpdateLastSavedTimeProperty
            Console.WriteLine("\nUpdated Last Saved Time Property successfully.");
        }

        public static void SetMeasureUnitForODT(string dataDir)
        {
            // ExStart:SetMeasureUnitForODT  
            //Load the Word document
            Document doc = new Document(dataDir + @"Document.doc");

            //Open Office uses centimeters when specifying lengths, widths and other measurable formatting 
            //and content properties in documents whereas MS Office uses inches. 

            OdtSaveOptions saveOptions = new OdtSaveOptions();
            saveOptions.MeasureUnit = OdtSaveMeasureUnit.Inches;

            //Save the document into ODT
            doc.Save(dataDir + "MeasureUnit_out.odt", saveOptions);
            // ExEnd:SetMeasureUnitForODT
            Console.WriteLine("\nSet MeasureUnit for ODT successfully.\nFile saved at " + dataDir);
        }
    }
}
