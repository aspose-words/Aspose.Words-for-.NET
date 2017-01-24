using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.IO;

namespace Simple_Bar_Graph
{
    class Program
    {
        // Simple bar graph 
        static void Main(string[] args)
        {
            // Check for license and apply if exists
            string licenseFile = AppDomain.CurrentDomain.BaseDirectory + "Aspose.Words.lic";
            if (File.Exists(licenseFile))
            {
                // Apply Aspose.Words API License
				Aspose.Words.License license = new Aspose.Words.License();
				// Place license file in Bin/Debug/Folder
				license.SetLicense("Aspose.Words.lic");
            }

			//createing new document
			Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write text on the document 
            builder.Write("Simple Bar graph using Aspose.Words \t");

            //select the chart type (here chartType is bar) 
            Shape shape1 = builder.InsertChart(ChartType.Bar, 432, 252);

            // save the document in the given path
            doc.Save("SimpleBarGraph.doc");
        }
    }
}
