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
            DocumentBuilder builder = new DocumentBuilder(doc);

            doc.Lists.Add(ListTemplate.NumberDefault);

            List list = doc.Lists[0];
            list.IsRestartAtEachSection = true;

            // The "IsRestartAtEachSection" property will only be applicable when
            // the document's OOXML compliance level is to a standard that is newer than "OoxmlComplianceCore.Ecma376".
            OoxmlSaveOptions options = new OoxmlSaveOptions
            {
                Compliance = OoxmlCompliance.Iso29500_2008_Transitional
            };

            builder.ListFormat.List = list;

            builder.Writeln("List item 1");
            builder.Writeln("List item 2");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("List item 3");
            builder.Writeln("List item 4");

            doc.Save(ArtifactsDir + "OoxmlSaveOptions.RestartingDocumentList.docx", options);
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