using Aspose.Words.Reporting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class ChartSeries
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LINQ();

            SetChartSeriesNameDynamically(dataDir);
            Console.WriteLine("\nChart template document is populated with the data.\nFile saved at " + dataDir);
        }

        public static void SetChartSeriesNameDynamically(string dataDir)
        {
            // ExStart:SetChartSeriesNameDynamically
            List<PointData> data = new List<PointData>()
            {
                new PointData { Time = "12:00:00 AM", Flow = 10, Rainfall = 2 },
                new PointData { Time = "01:00:00 AM", Flow = 15, Rainfall = 4 },
                new PointData { Time = "02:00:00 AM", Flow = 23, Rainfall = 7 }
            };

                        List<string> seriesNames = new List<string>
            {
                "Flow",
                "Rainfall"
            };

            Document doc = new Document(dataDir + "ChartTemplate.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, new object[] { data, seriesNames }, new string[] { "data", "seriesNames" });

            doc.Save(dataDir + "ChartTemplate_Out.docx");
            // ExEnd:SetChartSeriesNameDynamically
        }

    }
    // ExStart:PointDataClass
    public class PointData
    {
        public string Time { get; set; }
        public int Flow { get; set; }
        public int Rainfall { get; set; }
    }
    // ExEnd:PointDataClass
}
