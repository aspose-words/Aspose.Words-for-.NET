using Aspose.Words.Lists;
using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Lists
{
    class WorkingWithList
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithList();

            SetRestartAtEachSection(dataDir);
        }

        public static void SetRestartAtEachSection(String dataDir)
        {
            // ExStart:SetRestartAtEachSection
            Document doc = new Document(dataDir);

            doc.Lists.Add(ListTemplate.NumberDefault);

            List list = doc.Lists[0];

            // Set true to specify that the list has to be restarted at each section.
            list.IsRestartAtEachSection = true;

            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ListFormat.List = list;

            for (int i = 1; i < 45; i++)
            {
                builder.Writeln(String.Format("List Item {0}", i));

                // Insert section break.
                if (i == 15)
                    builder.InsertBreak(BreakType.SectionBreakNewPage);
            }

            // IsRestartAtEachSection will be written only if compliance is higher then OoxmlComplianceCore.Ecma376
            OoxmlSaveOptions options = new OoxmlSaveOptions();
            options.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;

            dataDir = dataDir + "RestartAtEachSection_out.docx";

            // Save the document to disk.
            doc.Save(dataDir, options);
            // ExEnd:SetRestartAtEachSection
            Console.WriteLine("\nDocument is saved successfully.");
        }
    }
}
