using Aspose.Words;
using Aspose.Words.Reporting;
using NUnit.Framework;

namespace DocsExamples.LINQ_Reporting_Engine
{
    public class BuildOptions : DocsExamplesBase
    {
        [Test]
        public void RemoveEmptyParagraphs()
        {
            //ExStart:RemoveEmptyParagraphs
            Document doc = new Document(MyDir + "Reporting engine template - Remove empty paragraphs.docx");

            ReportingEngine engine = new ReportingEngine { Options = ReportBuildOptions.RemoveEmptyParagraphs };
            engine.BuildReport(doc, Helpers.Common.GetManagers(), "Managers");

            doc.Save(ArtifactsDir + "ReportingEngine.RemoveEmptyParagraphs.docx");
            //ExEnd:RemoveEmptyParagraphs
        }
    }
}