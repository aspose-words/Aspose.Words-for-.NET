using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace Simple_Bar_Graph
{
    class Program
    {
        // Simple bar graph 
        static void Main(string[] args)
        {
            
           //file location
            string fileDir = "../../data/SimpleBarGraph.doc";
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
