using Aspose.Words;
using NUnit.Framework;
using System;

namespace DocsExamples.Programming_with_Documents
{
    internal class WorkingWithSection : DocsExamplesBase
    {
        [Test]
        public void AppendSectionContent()
        {
            //ExStart:NewTestAppendSectionContent
            //GistId:f44c5f27ccf595ae98813e7588d4e2d3
			//ReleaseNotesFunc:Added an ability to set chart axis title
			//ReleaseVersion:23.9
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Section 5");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write("Section 2");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write("Section 3");

            Section section = doc.Sections[2];

            // Insert the contents of the first section to the beginning of the third section.
            Section sectionToPrepend = doc.Sections[0];
            section.PrependContent(sectionToPrepend);

            // Insert the contents of the second section to the end of the third section.
            Section sectionToAppend = doc.Sections[1];
            section.AppendContent(sectionToAppend);
            //ExEnd:NewTestAppendSectionContent
        }
		
		[Test]
        public void AppendSectionContent()
        {
            //ExStart:NewTestAppendSectionContent2
            //GistId:f44c5f27ccf595ae98813e7588d4e2d3
			//RNDecs:Added public property MarkdownSaveOptions.ImagesFolderAlias
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Section 5");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write("Section 2");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write("Section 3");

            Section section = doc.Sections[2];

            // Insert the contents of the first section to the beginning of the third section.
            Section sectionToPrepend = doc.Sections[0];
            section.PrependContent(sectionToPrepend);

            // Insert the contents of the second section to the end of the third section.
            Section sectionToAppend = doc.Sections[1];
            section.AppendContent(sectionToAppend);
            //ExEnd:NewTestAppendSectionContent2
        }
    }
}