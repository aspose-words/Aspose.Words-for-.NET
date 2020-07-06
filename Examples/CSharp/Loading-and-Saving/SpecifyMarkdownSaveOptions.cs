using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Loading_and_Saving
{
    class SpecifyMarkdownSaveOptions
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            SaveAsMD(dataDir);
            ExportIntoMarkdownWithTableContentAlignment(dataDir);
        }

        private static void SaveAsMD(string dataDir)
        {
            // ExStart:SaveAsMD
            DocumentBuilder builder = new DocumentBuilder();
            builder.Writeln("Some text!");

            // specify MarkDownSaveOptions
            MarkdownSaveOptions saveOptions = (MarkdownSaveOptions)SaveOptions.CreateSaveOptions(SaveFormat.Markdown);
            
            builder.Document.Save(dataDir + "TestDocument.md", saveOptions);
            // ExEnd:SaveAsMD
        }

        private static void ExportIntoMarkdownWithTableContentAlignment(string dataDir)
        {
            // ExStart:ExportIntoMarkdownWithTableContentAlignment
            DocumentBuilder builder = new DocumentBuilder();

            // Create a new table with two cells.
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            builder.Write("Cell1");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Write("Cell2");

            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
            // Makes all paragraphs inside table to be aligned to Left. 
            saveOptions.TableContentAlignment = TableContentAlignment.Left;
            builder.Document.Save(dataDir + "left.md", saveOptions);

            // Makes all paragraphs inside table to be aligned to Right. 
            saveOptions.TableContentAlignment = TableContentAlignment.Right;
            builder.Document.Save(dataDir + "right.md", saveOptions);

            // Makes all paragraphs inside table to be aligned to Center. 
            saveOptions.TableContentAlignment = TableContentAlignment.Center;
            builder.Document.Save(dataDir + "center.md", saveOptions);

            // Makes all paragraphs inside table to be aligned automatically.
            // The alignment in this case will be taken from the first paragraph in corresponding table column.
            saveOptions.TableContentAlignment = TableContentAlignment.Auto;
            builder.Document.Save(dataDir + "auto.md", saveOptions);
            // ExEnd:ExportIntoMarkdownWithTableContentAlignment
        }
    }
}
