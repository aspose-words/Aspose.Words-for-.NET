using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Save_Options
{
    public class WorkingWithTxtSaveOptions : DocsExamplesBase
    {
        [Test]
        public void AddBidiMarks()
        {
            //ExStart:AddBidiMarks
            //GistId:ddafc3430967fb4f4f70085fa577d01a
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");
            builder.ParagraphFormat.Bidi = true;
            builder.Writeln("שלום עולם!");
            builder.Writeln("مرحبا بالعالم!");

            TxtSaveOptions saveOptions = new TxtSaveOptions { AddBidiMarks = true };

            doc.Save(ArtifactsDir + "WorkingWithTxtSaveOptions.AddBidiMarks.txt", saveOptions);
            //ExEnd:AddBidiMarks
        }

        [Test]
        public void UseTabForListIndentation()
        {
            //ExStart:UseTabForListIndentation
            //GistId:ddafc3430967fb4f4f70085fa577d01a
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a list with three levels of indentation.
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("Item 1");
            builder.ListFormat.ListIndent();
            builder.Writeln("Item 2");
            builder.ListFormat.ListIndent(); 
            builder.Write("Item 3");

            TxtSaveOptions saveOptions = new TxtSaveOptions();
            saveOptions.ListIndentation.Count = 1;
            saveOptions.ListIndentation.Character = '\t';

            doc.Save(ArtifactsDir + "WorkingWithTxtSaveOptions.UseTabForListIndentation.txt", saveOptions);
            //ExEnd:UseTabForListIndentation
        }

        [Test]
        public void UseSpaceForListIndentation()
        {
            //ExStart:UseSpaceForListIndentation
            //GistId:ddafc3430967fb4f4f70085fa577d01a
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a list with three levels of indentation.
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("Item 1");
            builder.ListFormat.ListIndent();
            builder.Writeln("Item 2");
            builder.ListFormat.ListIndent(); 
            builder.Write("Item 3");

            TxtSaveOptions saveOptions = new TxtSaveOptions();
            saveOptions.ListIndentation.Count = 3;
            saveOptions.ListIndentation.Character = ' ';

            doc.Save(ArtifactsDir + "WorkingWithTxtSaveOptions.UseSpaceForListIndentation.txt", saveOptions);
            //ExEnd:UseSpaceForListIndentation
        }
    }
}