using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;
using NUnit.Framework;
using List = Aspose.Words.Lists.List;

namespace DocsExamples.Programming_with_Documents
{
    internal class WorkingWithList : DocsExamplesBase
    {
        [Test]
        public void RestartListAtEachSection()
        {
            //ExStart:RestartListAtEachSection
            Document doc = new Document();
            
            doc.Lists.Add(ListTemplate.NumberDefault);

            List list = doc.Lists[0];
            list.IsRestartAtEachSection = true;

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.ListFormat.List = list;

            for (int i = 1; i < 45; i++)
            {
                builder.Writeln($"List Item {i}");

                if (i == 15)
                    builder.InsertBreak(BreakType.SectionBreakNewPage);
            }

            // IsRestartAtEachSection will be written only if compliance is higher then OoxmlComplianceCore.Ecma376.
            OoxmlSaveOptions options = new OoxmlSaveOptions { Compliance = OoxmlCompliance.Iso29500_2008_Transitional };

            doc.Save(ArtifactsDir + "WorkingWithList.RestartListAtEachSection.docx", options);
            //ExEnd:RestartListAtEachSection
        }

        [Test]
        public void SpecifyListLevel()
        {
            //ExStart:SpecifyListLevel
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a numbered list based on one of the Microsoft Word list templates
            // and apply it to the document builder's current paragraph.
            builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberArabicDot);

            // There are nine levels in this list, let's try them all.
            for (int i = 0; i < 9; i++)
            {
                builder.ListFormat.ListLevelNumber = i;
                builder.Writeln("Level " + i);
            }

            // Create a bulleted list based on one of the Microsoft Word list templates
            // and apply it to the document builder's current paragraph.
            builder.ListFormat.List = doc.Lists.Add(ListTemplate.BulletDiamonds);

            for (int i = 0; i < 9; i++)
            {
                builder.ListFormat.ListLevelNumber = i;
                builder.Writeln("Level " + i);
            }

            // This is a way to stop list formatting.
            builder.ListFormat.List = null;

            builder.Document.Save(ArtifactsDir + "WorkingWithList.SpecifyListLevel.docx");
            //ExEnd:SpecifyListLevel
        }

        [Test]
        public void RestartListNumber()
        {
            //ExStart:RestartListNumber
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a list based on a template.
            List list1 = doc.Lists.Add(ListTemplate.NumberArabicParenthesis);
            list1.ListLevels[0].Font.Color = Color.Red;
            list1.ListLevels[0].Alignment = ListLevelAlignment.Right;

            builder.Writeln("List 1 starts below:");
            builder.ListFormat.List = list1;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();

            // To reuse the first list, we need to restart numbering by creating a copy of the original list formatting.
            List list2 = doc.Lists.AddCopy(list1);

            // We can modify the new list in any way, including setting a new start number.
            list2.ListLevels[0].StartAt = 10;

            builder.Writeln("List 2 starts below:");
            builder.ListFormat.List = list2;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();

            builder.Document.Save(ArtifactsDir + "WorkingWithList.RestartListNumber.docx");
            //ExEnd:RestartListNumber
        }
    }
}