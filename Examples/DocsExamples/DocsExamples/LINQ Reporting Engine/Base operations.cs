using System.Collections.Generic;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Reporting;
using DocsExamples.LINQ_Reporting_Engine.Helpers.Data_Source_Objects;
using NUnit.Framework;

namespace DocsExamples.LINQ_Reporting_Engine
{
    public class BaseOperations : DocsExamplesBase
    {
        [Test]
        public void HelloWorld()
        {
            //ExStart:HelloWorld
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Write("<<[sender.Name]>> says: <<[sender.Message]>>");

            Sender sender = new Sender { Name = "LINQ Reporting Engine", Message = "Hello World" };

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, sender, "sender");

            doc.Save(ArtifactsDir + "ReportingEngine.HelloWorld.docx");
            //ExEnd:HelloWorld
        }

        [Test]
        public void SingleRow()
        {
            //ExStart:SingleRow
            Document doc = new Document(MyDir + "Reporting engine template - Table row.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetManagers(), "Managers");

            doc.Save(ArtifactsDir + "ReportingEngine.SingleRow.docx");
            //ExEnd:SingleRow
        }

        [Test]
        public void CommonMasterDetail()
        {
            //ExStart:CommonMasterDetail
            Document doc = new Document(MyDir + "Reporting engine template - Common master detail.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetManagers(), "managers");

            doc.Save(ArtifactsDir + "ReportingEngine.CommonMasterDetail.docx");
            //ExEnd:CommonMasterDetail
        }

        [Test]
        public void ConditionalBlocks()
        {
            //ExStart:ConditionalBlocks
            Document doc = new Document(MyDir + "Reporting engine template - Table row conditional blocks.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetClients(), "clients");

            doc.Save(ArtifactsDir + "ReportingEngine.ConditionalBlock.docx");
            //ExEnd:ConditionalBlocks
        }

        [Test]
        public void SettingBackgroundColor()
        {
            //ExStart:SettingBackgroundColor
            Document doc = new Document(MyDir + "Reporting engine template - Background color.docx");

            List<BackgroundColor> colors = new List<BackgroundColor>
            {
                new BackgroundColor {Name = "Black", Color = Color.Black},
                new BackgroundColor {Name = "Red", Color = Color.FromArgb(255, 0, 0)},
                new BackgroundColor {Name = "Empty", Color = Color.Empty}
            };

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, colors, "Colors");

            doc.Save(ArtifactsDir + "ReportingEngine.BackColor.docx");
            //ExEnd:SettingBackgroundColor
        }
    }
}