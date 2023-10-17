using Aspose.Words;
using NUnit.Framework;
using System.Data;

namespace PluginsExamples
{
    public class MailMergePlugin : PluginsExamplesBase
    {
        [Test]
        public void SimpleMailMerge()
        {
            //ExStart:SimpleMailMerge
            //GistId:bca6f72734f250c63a1285df42bd2498
            var doc = new Document(MyDir + "Mail merge template.docx");
            doc.MailMerge.Execute(new[] { "FieldName" }, new object[] { "Value" });
            doc.Save(ArtifactsDir + "MailMergePlugin.SimpleMailMerge.docx");
            //ExEnd:SimpleMailMerge
        }

        [Test]
        public void MailMergeWithXml()
        {
            //ExStart:MailMergeWithXml
            //GistId:bca6f72734f250c63a1285df42bd2498
            DataSet data = new DataSet();
            data.ReadXml(MyDir + "Mail merge data - Orders.xml");

            Document doc = new Document(MyDir + "Mail merge destinations - Invoice.docx");
            // Trim trailing and leading whitespaces mail merge values.
            doc.MailMerge.TrimWhitespaces = false;
            doc.MailMerge.ExecuteWithRegions(data);

            doc.Save(ArtifactsDir + "MailMergePlugin.MailMergeWithXml.docx");
            //ExEnd:MailMergeWithXml
        }
    }
}
