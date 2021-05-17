using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Save_Options
{
    public class WorkingWithMarkdownSaveOptions : DocsExamplesBase
    {
        [Test]
        public void ExportIntoMarkdownWithTableContentAlignment()
        {
            //ExStart:ExportIntoMarkdownWithTableContentAlignment
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            builder.Write("Cell1");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Write("Cell2");

            // Makes all paragraphs inside the table to be aligned.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                TableContentAlignment = TableContentAlignment.Left
            };
            doc.Save(ArtifactsDir + "WorkingWithMarkdownSaveOptions.LeftTableContentAlignment.md", saveOptions);

            saveOptions.TableContentAlignment = TableContentAlignment.Right;
            doc.Save(ArtifactsDir + "WorkingWithMarkdownSaveOptions.RightTableContentAlignment.md", saveOptions);

            saveOptions.TableContentAlignment = TableContentAlignment.Center;
            doc.Save(ArtifactsDir + "WorkingWithMarkdownSaveOptions.CenterTableContentAlignment.md", saveOptions);

            // The alignment in this case will be taken from the first paragraph in corresponding table column.
            saveOptions.TableContentAlignment = TableContentAlignment.Auto;
            doc.Save(ArtifactsDir + "WorkingWithMarkdownSaveOptions.AutoTableContentAlignment.md", saveOptions);
            //ExEnd:ExportIntoMarkdownWithTableContentAlignment
        }

        [Test]
        public void SetImagesFolder()
        {
            //ExStart:SetImagesFolder
            Document doc = new Document(MyDir + "Image bullet points.docx");

            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions { ImagesFolder = ArtifactsDir + "Images" };

            using (MemoryStream stream = new MemoryStream())
                doc.Save(stream, saveOptions);
            //ExEnd:SetImagesFolder
        }
    }
}
