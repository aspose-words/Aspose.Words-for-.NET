using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.IO;

namespace Simple_Bar_Graph
{
    class Program
    {
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

			Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Simple Bar graph using Aspose.Words \t");

            Shape shape1 = builder.InsertChart(ChartType.Bar, 432, 252);

            doc.Save("Simple_Bar_Graph.docx");
        }
    }
}
