using Aspose.Words.Lists;
using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Drawing;
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
            RestartListNumber(dataDir);
            SpecifyListLevel(dataDir);
            SetRestartAtEachSection(dataDir);
        }

        public static void SetRestartAtEachSection(String dataDir)
        {
            // ExStart:SetRestartAtEachSection
            Document doc = new Document();

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
            Console.WriteLine("\nDocument is saved successfully.\nFile saved at " + dataDir);
        }

        public static void SpecifyListLevel(String dataDir)
        {
            // ExStart:SpecifyListLevel
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a numbered list based on one of the Microsoft Word list templates and
            // apply it to the current paragraph in the document builder.
            builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberArabicDot);

            // There are 9 levels in this list, lets try them all.
            for (int i = 0; i < 9; i++)
            {
                builder.ListFormat.ListLevelNumber = i;
                builder.Writeln("Level " + i);
            }


            // Create a bulleted list based on one of the Microsoft Word list templates
            // and apply it to the current paragraph in the document builder.
            builder.ListFormat.List = doc.Lists.Add(ListTemplate.BulletDiamonds);

            // There are 9 levels in this list, lets try them all.
            for (int i = 0; i < 9; i++)
            {
                builder.ListFormat.ListLevelNumber = i;
                builder.Writeln("Level " + i);
            }

            // This is a way to stop list formatting. 
            builder.ListFormat.List = null;

            dataDir = dataDir + "Lists.SpecifyListLevel Out.doc";

            // Save the document to disk.
            builder.Document.Save(dataDir);
            // ExEnd:SpecifyListLevel
            Console.WriteLine("\nDocument is saved successfully.\nFile saved at " + dataDir);
        }


        public static void RestartListNumber(String dataDir)
        {
            // ExStart:RestartListNumber
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a list based on a template.
            Aspose.Words.Lists.List list1 = doc.Lists.Add(ListTemplate.NumberArabicParenthesis);
            // Modify the formatting of the list.
            list1.ListLevels[0].Font.Color = Color.Red;
            list1.ListLevels[0].Alignment = ListLevelAlignment.Right;

            builder.Writeln("List 1 starts below:");
            // Use the first list in the document for a while.
            builder.ListFormat.List = list1;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();

            // Now I want to reuse the first list, but need to restart numbering.
            // This should be done by creating a copy of the original list formatting.
            Aspose.Words.Lists.List list2 = doc.Lists.AddCopy(list1);

            // We can modify the new list in any way. Including setting new start number.
            list2.ListLevels[0].StartAt = 10;

            // Use the second list in the document.
            builder.Writeln("List 2 starts below:");
            builder.ListFormat.List = list2;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();

            dataDir = dataDir + "Lists.RestartNumberingUsingListCopy Out.doc";

            // Save the document to disk.
            builder.Document.Save(dataDir);
            // ExEnd:RestartListNumber
            Console.WriteLine("\nDocument is saved successfully.\nFile saved at " + dataDir);
        }
    }
}
