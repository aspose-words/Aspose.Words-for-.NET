﻿// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Drawing;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;
using NUnit.Framework;
using List = Aspose.Words.Lists.List;

namespace ApiExamples
{
    [TestFixture]
    public class ExLists : ApiExampleBase
    {
        [Test]
        public void ApplyDefaultBulletsAndNumbers()
        {
            //ExStart
            //ExFor:DocumentBuilder.ListFormat
            //ExFor:ListFormat.ApplyNumberDefault
            //ExFor:ListFormat.ApplyBulletDefault
            //ExFor:ListFormat.ListIndent
            //ExFor:ListFormat.ListOutdent
            //ExFor:ListFormat.RemoveNumbers
            //ExFor:ListFormat.ListLevelNumber
            //ExSummary:Shows how to create bulleted and numbered lists.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Aspose.Words main advantages are:");

            // A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
            // We can create nested lists by increasing the indent level. 
            // We can begin and end a list by using a document builder's "ListFormat" property. 
            // Each paragraph that we add between a list's start and the end will become an item in the list.
            // Below are two types of lists that we can create with a document builder.
            // 1 -  A bulleted list:
            // This list will apply an indent and a bullet symbol ("•") before each paragraph.
            builder.ListFormat.ApplyBulletDefault();
            builder.Writeln("Great performance");
            builder.Writeln("High reliability");
            builder.Writeln("Quality code and working");
            builder.Writeln("Wide variety of features");
            builder.Writeln("Easy to understand API");

            // End the bulleted list.
            builder.ListFormat.RemoveNumbers();

            builder.InsertBreak(BreakType.ParagraphBreak);
            builder.Writeln("Aspose.Words allows:");

            // 2 -  A numbered list:
            // Numbered lists create a logical order for their paragraphs by numbering each item.
            builder.ListFormat.ApplyNumberDefault();

            // This paragraph is the first item. The first item of a numbered list will have a "1." as its list item symbol.
            builder.Writeln("Opening documents from different formats:");

            Assert.That(builder.ListFormat.ListLevelNumber, Is.EqualTo(0));

            // Call the "ListIndent" method to increase the current list level,
            // which will start a new self-contained list, with a deeper indent, at the current item of the first list level.
            builder.ListFormat.ListIndent();

            Assert.That(builder.ListFormat.ListLevelNumber, Is.EqualTo(1));

            // These are the first three list items of the second list level, which will maintain a count
            // independent of the count of the first list level. According to the current list format,
            // they will have symbols of "a.", "b.", and "c.".
            builder.Writeln("DOC");
            builder.Writeln("PDF");
            builder.Writeln("HTML");

            // Call the "ListOutdent" method to return to the previous list level.
            builder.ListFormat.ListOutdent();

            Assert.That(builder.ListFormat.ListLevelNumber, Is.EqualTo(0));

            // These two paragraphs will continue the count of the first list level.
            // These items will have symbols of "2.", and "3."
            builder.Writeln("Processing documents");
            builder.Writeln("Saving documents in different formats:");

            // If we increase the list level to a level that we have added items to previously,
            // the nested list will be separate from the previous, and its numbering will start from the beginning. 
            // These list items will have symbols of "a.", "b.", "c.", "d.", and "e".
            builder.ListFormat.ListIndent();
            builder.Writeln("DOC");
            builder.Writeln("PDF");
            builder.Writeln("HTML");
            builder.Writeln("MHTML");
            builder.Writeln("Plain text");

            // Outdent the list level again.
            builder.ListFormat.ListOutdent();
            builder.Writeln("Doing many other things!");

            // End the numbered list.
            builder.ListFormat.RemoveNumbers();

            doc.Save(ArtifactsDir + "Lists.ApplyDefaultBulletsAndNumbers.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Lists.ApplyDefaultBulletsAndNumbers.docx");

            TestUtil.VerifyListLevel("\0.", 18.0d, NumberStyle.Arabic, doc.Lists[1].ListLevels[0]);
            TestUtil.VerifyListLevel("\u0001.", 54.0d, NumberStyle.LowercaseLetter, doc.Lists[1].ListLevels[1]);
            TestUtil.VerifyListLevel("\uf0b7", 18.0d, NumberStyle.Bullet, doc.Lists[0].ListLevels[0]);
        }

        [Test]
        public void SpecifyListLevel()
        {
            //ExStart
            //ExFor:ListCollection
            //ExFor:List
            //ExFor:ListFormat
            //ExFor:ListFormat.IsListItem
            //ExFor:ListFormat.ListLevelNumber
            //ExFor:ListFormat.List
            //ExFor:ListTemplate
            //ExFor:DocumentBase.Lists
            //ExFor:ListCollection.Add(ListTemplate)
            //ExSummary:Shows how to work with list levels.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Assert.That(builder.ListFormat.IsListItem, Is.False);

            // A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
            // We can create nested lists by increasing the indent level. 
            // We can begin and end a list by using a document builder's "ListFormat" property. 
            // Each paragraph that we add between a list's start and the end will become an item in the list.
            // Below are two types of lists that we can create using a document builder.
            // 1 -  A numbered list:
            // Numbered lists create a logical order for their paragraphs by numbering each item.
            builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

            Assert.That(builder.ListFormat.IsListItem, Is.True);

            // By setting the "ListLevelNumber" property, we can increase the list level
            // to begin a self-contained sub-list at the current list item.
            // The Microsoft Word list template called "NumberDefault" uses numbers to create list levels for the first list level.
            // Deeper list levels use letters and lowercase Roman numerals. 
            for (int i = 0; i < 9; i++)
            {
                builder.ListFormat.ListLevelNumber = i;
                builder.Writeln("Level " + i);
            }

            // 2 -  A bulleted list:
            // This list will apply an indent and a bullet symbol ("•") before each paragraph.
            // Deeper levels of this list will use different symbols, such as "■" and "○".
            builder.ListFormat.List = doc.Lists.Add(ListTemplate.BulletDefault);

            for (int i = 0; i < 9; i++)
            {
                builder.ListFormat.ListLevelNumber = i;
                builder.Writeln("Level " + i);
            }

            // We can disable list formatting to not format any subsequent paragraphs as lists by un-setting the "List" flag.
            builder.ListFormat.List = null;

            Assert.That(builder.ListFormat.IsListItem, Is.False);

            doc.Save(ArtifactsDir + "Lists.SpecifyListLevel.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Lists.SpecifyListLevel.docx");

            TestUtil.VerifyListLevel("\0.", 18.0d, NumberStyle.Arabic, doc.Lists[0].ListLevels[0]);
        }

        [Test]
        public void NestedLists()
        {
            //ExStart
            //ExFor:ListFormat.List
            //ExFor:ParagraphFormat.ClearFormatting
            //ExFor:ParagraphFormat.DropCapPosition
            //ExFor:ParagraphFormat.IsListItem
            //ExFor:Paragraph.IsListItem
            //ExSummary:Shows how to nest a list inside another list.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
            // We can create nested lists by increasing the indent level. 
            // We can begin and end a list by using a document builder's "ListFormat" property. 
            // Each paragraph that we add between a list's start and the end will become an item in the list.
            // Create an outline list for the headings.
            List outlineList = doc.Lists.Add(ListTemplate.OutlineNumbers);
            builder.ListFormat.List = outlineList;
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("This is my Chapter 1");

            // Create a numbered list.
            List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);
            builder.ListFormat.List = numberedList;
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Numbered list item 1.");

            // Every paragraph that comprises a list will have this flag.
            Assert.That(builder.CurrentParagraph.IsListItem, Is.True);
            Assert.That(builder.ParagraphFormat.IsListItem, Is.True);

            // Create a bulleted list.
            List bulletedList = doc.Lists.Add(ListTemplate.BulletDefault);
            builder.ListFormat.List = bulletedList;
            builder.ParagraphFormat.LeftIndent = 72;
            builder.Writeln("Bulleted list item 1.");
            builder.Writeln("Bulleted list item 2.");
            builder.ParagraphFormat.ClearFormatting();

            // Revert to the numbered list.
            builder.ListFormat.List = numberedList;
            builder.Writeln("Numbered list item 2.");
            builder.Writeln("Numbered list item 3.");

            // Revert to the outline list.
            builder.ListFormat.List = outlineList;
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("This is my Chapter 2");

            builder.ParagraphFormat.ClearFormatting();

            builder.Document.Save(ArtifactsDir + "Lists.NestedLists.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Lists.NestedLists.docx");

            TestUtil.VerifyListLevel("\0)", 0.0d, NumberStyle.Arabic, doc.Lists[0].ListLevels[0]);
            TestUtil.VerifyListLevel("\0.", 18.0d, NumberStyle.Arabic, doc.Lists[1].ListLevels[0]);
            TestUtil.VerifyListLevel("\uf0b7", 18.0d, NumberStyle.Bullet, doc.Lists[2].ListLevels[0]);
        }

        [Test]
        public void CreateCustomList()
        {
            //ExStart
            //ExFor:List
            //ExFor:List.ListLevels
            //ExFor:ListFormat.ListLevel
            //ExFor:ListLevelCollection
            //ExFor:ListLevelCollection.Item
            //ExFor:ListLevel
            //ExFor:ListLevel.Alignment
            //ExFor:ListLevel.Font
            //ExFor:ListLevel.NumberStyle
            //ExFor:ListLevel.StartAt
            //ExFor:ListLevel.TrailingCharacter
            //ExFor:ListLevelAlignment
            //ExFor:NumberStyle
            //ExFor:ListTrailingCharacter
            //ExFor:ListLevel.NumberFormat
            //ExFor:ListLevel.NumberPosition
            //ExFor:ListLevel.TextPosition
            //ExFor:ListLevel.TabPosition
            //ExSummary:Shows how to apply custom list formatting to paragraphs when using DocumentBuilder.
            Document doc = new Document();

            // A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
            // We can create nested lists by increasing the indent level. 
            // We can begin and end a list by using a document builder's "ListFormat" property. 
            // Each paragraph that we add between a list's start and the end will become an item in the list.
            // Create a list from a Microsoft Word template, and customize the first two of its list levels.
            List docList = doc.Lists.Add(ListTemplate.NumberDefault);

            ListLevel listLevel = docList.ListLevels[0];
            listLevel.Font.Color = Color.Red;
            listLevel.Font.Size = 24;
            listLevel.NumberStyle = NumberStyle.OrdinalText;
            listLevel.StartAt = 21;
            listLevel.NumberFormat = "\x0000";

            listLevel.NumberPosition = -36;
            listLevel.TextPosition = 144;
            listLevel.TabPosition = 144;

            listLevel = docList.ListLevels[1];
            listLevel.Alignment = ListLevelAlignment.Right;
            listLevel.NumberStyle = NumberStyle.Bullet;
            listLevel.Font.Name = "Wingdings";
            listLevel.Font.Color = Color.Blue;
            listLevel.Font.Size = 24;

            // This NumberFormat value will create star-shaped bullet list symbols.
            listLevel.NumberFormat = "\xf0af";
            listLevel.TrailingCharacter = ListTrailingCharacter.Space;
            listLevel.NumberPosition = 144;

            // Create paragraphs and apply both list levels of our custom list formatting to them.
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ListFormat.List = docList;
            builder.Writeln("The quick brown fox...");
            builder.Writeln("The quick brown fox...");

            builder.ListFormat.ListIndent();
            builder.Writeln("jumped over the lazy dog.");
            builder.Writeln("jumped over the lazy dog.");

            builder.ListFormat.ListOutdent();
            builder.Writeln("The quick brown fox...");

            builder.ListFormat.RemoveNumbers();

            builder.Document.Save(ArtifactsDir + "Lists.CreateCustomList.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Lists.CreateCustomList.docx");

            listLevel = doc.Lists[0].ListLevels[0];

            TestUtil.VerifyListLevel("\0", -36.0d, NumberStyle.OrdinalText, listLevel);
            Assert.That(listLevel.Font.Color.ToArgb(), Is.EqualTo(Color.Red.ToArgb()));
            Assert.That(listLevel.Font.Size, Is.EqualTo(24.0d));
            Assert.That(listLevel.StartAt, Is.EqualTo(21));

            listLevel = doc.Lists[0].ListLevels[1];

            TestUtil.VerifyListLevel("\xf0af", 144.0d, NumberStyle.Bullet, listLevel);
            Assert.That(listLevel.Font.Color.ToArgb(), Is.EqualTo(Color.Blue.ToArgb()));
            Assert.That(listLevel.Font.Size, Is.EqualTo(24.0d));
            Assert.That(listLevel.StartAt, Is.EqualTo(1));
            Assert.That(listLevel.TrailingCharacter, Is.EqualTo(ListTrailingCharacter.Space));
        }

        [Test]
        public void RestartNumberingUsingListCopy()
        {
            //ExStart
            //ExFor:List
            //ExFor:ListCollection
            //ExFor:ListCollection.Add(ListTemplate)
            //ExFor:ListCollection.AddCopy(List)
            //ExFor:ListLevel.StartAt
            //ExFor:ListTemplate
            //ExSummary:Shows how to restart numbering in a list by copying a list.
            Document doc = new Document();

            // A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
            // We can create nested lists by increasing the indent level. 
            // We can begin and end a list by using a document builder's "ListFormat" property. 
            // Each paragraph that we add between a list's start and the end will become an item in the list.
            // Create a list from a Microsoft Word template, and customize its first list level.
            List list1 = doc.Lists.Add(ListTemplate.NumberArabicParenthesis);
            list1.ListLevels[0].Font.Color = Color.Red;
            list1.ListLevels[0].Alignment = ListLevelAlignment.Right;

            // Apply our list to some paragraphs.
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("List 1 starts below:");
            builder.ListFormat.List = list1;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();

            // We can add a copy of an existing list to the document's list collection
            // to create a similar list without making changes to the original.
            List list2 = doc.Lists.AddCopy(list1);
            list2.ListLevels[0].Font.Color = Color.Blue;
            list2.ListLevels[0].StartAt = 10;

            // Apply the second list to new paragraphs.
            builder.Writeln("List 2 starts below:");
            builder.ListFormat.List = list2;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();

            doc.Save(ArtifactsDir + "Lists.RestartNumberingUsingListCopy.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Lists.RestartNumberingUsingListCopy.docx");

            list1 = doc.Lists[0];
            TestUtil.VerifyListLevel("\0)", 18.0d, NumberStyle.Arabic, list1.ListLevels[0]);
            Assert.That(list1.ListLevels[0].Font.Color.ToArgb(), Is.EqualTo(Color.Red.ToArgb()));
            Assert.That(list1.ListLevels[0].Font.Size, Is.EqualTo(10.0d));
            Assert.That(list1.ListLevels[0].StartAt, Is.EqualTo(1));

            list2 = doc.Lists[1];
            TestUtil.VerifyListLevel("\0)", 18.0d, NumberStyle.Arabic, list2.ListLevels[0]);
            Assert.That(list2.ListLevels[0].Font.Color.ToArgb(), Is.EqualTo(Color.Blue.ToArgb()));
            Assert.That(list2.ListLevels[0].Font.Size, Is.EqualTo(10.0d));
            Assert.That(list2.ListLevels[0].StartAt, Is.EqualTo(10));
        }

        [Test]
        public void CreateAndUseListStyle()
        {
            //ExStart
            //ExFor:StyleCollection.Add(StyleType,String)
            //ExFor:Style.List
            //ExFor:StyleType
            //ExFor:List.IsListStyleDefinition
            //ExFor:List.IsListStyleReference
            //ExFor:List.IsMultiLevel
            //ExFor:List.Style
            //ExFor:ListLevelCollection
            //ExFor:ListLevelCollection.Count
            //ExFor:ListLevelCollection.Item
            //ExFor:ListCollection.Add(Style)
            //ExSummary:Shows how to create a list style and use it in a document.
            Document doc = new Document();

            // A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
            // We can create nested lists by increasing the indent level. 
            // We can begin and end a list by using a document builder's "ListFormat" property. 
            // Each paragraph that we add between a list's start and the end will become an item in the list.
            // We can contain an entire List object within a style.
            Style listStyle = doc.Styles.Add(StyleType.List, "MyListStyle");

            List list1 = listStyle.List;

            Assert.That(list1.IsListStyleDefinition, Is.True);
            Assert.That(list1.IsListStyleReference, Is.False);
            Assert.That(list1.IsMultiLevel, Is.True);
            Assert.That(list1.Style, Is.EqualTo(listStyle));

            // Change the appearance of all list levels in our list.
            foreach (ListLevel level in list1.ListLevels)
            {
                level.Font.Name = "Verdana";
                level.Font.Color = Color.Blue;
                level.Font.Bold = true;
            }

            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Using list style first time:");

            // Create another list from a list within a style.
            List list2 = doc.Lists.Add(listStyle);

            Assert.That(list2.IsListStyleDefinition, Is.False);
            Assert.That(list2.IsListStyleReference, Is.True);
            Assert.That(list2.Style, Is.EqualTo(listStyle));

            // Add some list items that our list will format.
            builder.ListFormat.List = list2;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();

            builder.Writeln("Using list style second time:");

            // Create and apply another list based on the list style.
            List list3 = doc.Lists.Add(listStyle);
            builder.ListFormat.List = list3;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();

            builder.Document.Save(ArtifactsDir + "Lists.CreateAndUseListStyle.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Lists.CreateAndUseListStyle.docx");

            list1 = doc.Lists[0];

            TestUtil.VerifyListLevel("\0.", 18.0d, NumberStyle.Arabic, list1.ListLevels[0]);
            Assert.That(list1.IsListStyleDefinition, Is.True);
            Assert.That(list1.IsListStyleReference, Is.False);
            Assert.That(list1.IsMultiLevel, Is.True);
            Assert.That(list1.ListLevels[0].Font.Color.ToArgb(), Is.EqualTo(Color.Blue.ToArgb()));
            Assert.That(list1.ListLevels[0].Font.Name, Is.EqualTo("Verdana"));
            Assert.That(list1.ListLevels[0].Font.Bold, Is.True);

            list2 = doc.Lists[1];

            TestUtil.VerifyListLevel("\0.", 18.0d, NumberStyle.Arabic, list2.ListLevels[0]);
            Assert.That(list2.IsListStyleDefinition, Is.False);
            Assert.That(list2.IsListStyleReference, Is.True);
            Assert.That(list2.IsMultiLevel, Is.True);

            list3 = doc.Lists[2];

            TestUtil.VerifyListLevel("\0.", 18.0d, NumberStyle.Arabic, list3.ListLevels[0]);
            Assert.That(list3.IsListStyleDefinition, Is.False);
            Assert.That(list3.IsListStyleReference, Is.True);
            Assert.That(list3.IsMultiLevel, Is.True);
        }

        [Test]
        public void DetectBulletedParagraphs()
        {
            //ExStart
            //ExFor:Paragraph.ListFormat
            //ExFor:ListFormat.IsListItem
            //ExFor:CompositeNode.GetText
            //ExFor:List.ListId
            //ExSummary:Shows how to output all paragraphs in a document that are list items.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("Numbered list item 1");
            builder.Writeln("Numbered list item 2");
            builder.Writeln("Numbered list item 3");
            builder.ListFormat.RemoveNumbers();

            builder.ListFormat.ApplyBulletDefault();
            builder.Writeln("Bulleted list item 1");
            builder.Writeln("Bulleted list item 2");
            builder.Writeln("Bulleted list item 3");
            builder.ListFormat.RemoveNumbers();

            NodeCollection paras = doc.GetChildNodes(NodeType.Paragraph, true);

            foreach (Paragraph para in paras.OfType<Paragraph>().Where(p => p.ListFormat.IsListItem).ToList())
            { 
                Console.WriteLine($"This paragraph belongs to list ID# {para.ListFormat.List.ListId}, number style \"{para.ListFormat.ListLevel.NumberStyle}\"");
                Console.WriteLine($"\t\"{para.GetText().Trim()}\"");
            }
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            paras = doc.GetChildNodes(NodeType.Paragraph, true);

            Assert.That(paras.Count(n => ((Paragraph)n).ListFormat.IsListItem), Is.EqualTo(6));
        }

        [Test]
        public void RemoveBulletsFromParagraphs()
        {
            //ExStart
            //ExFor:ListFormat.RemoveNumbers
            //ExSummary:Shows how to remove list formatting from all paragraphs in the main text of a section.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("Numbered list item 1");
            builder.Writeln("Numbered list item 2");
            builder.Writeln("Numbered list item 3");
            builder.ListFormat.RemoveNumbers();

            NodeCollection paras = doc.GetChildNodes(NodeType.Paragraph, true);
            Assert.That(paras.Count(n => ((Paragraph)n).ListFormat.IsListItem), Is.EqualTo(3));

            foreach (Paragraph paragraph in paras)
                paragraph.ListFormat.RemoveNumbers();

            Assert.That(paras.Count(n => ((Paragraph)n).ListFormat.IsListItem), Is.EqualTo(0));
            //ExEnd
        }

        [Test]
        public void ApplyExistingListToParagraphs()
        {
            //ExStart
            //ExFor:ListCollection.Item(Int32)
            //ExSummary:Shows how to apply list formatting of an existing list to a collection of paragraphs.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Paragraph 1");
            builder.Writeln("Paragraph 2");
            builder.Write("Paragraph 3");

            NodeCollection paras = doc.GetChildNodes(NodeType.Paragraph, true);

            Assert.That(paras.Count(n => ((Paragraph)n).ListFormat.IsListItem), Is.EqualTo(0));

            doc.Lists.Add(ListTemplate.NumberDefault);
            List docList = doc.Lists[0];

            foreach (Paragraph paragraph in paras.OfType<Paragraph>())
            {
                paragraph.ListFormat.List = docList;
                paragraph.ListFormat.ListLevelNumber = 2;
            }

            Assert.That(paras.Count(n => ((Paragraph)n).ListFormat.IsListItem), Is.EqualTo(3));
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            paras = doc.GetChildNodes(NodeType.Paragraph, true);

            Assert.That(paras.Count(n => ((Paragraph)n).ListFormat.IsListItem), Is.EqualTo(3));
            Assert.That(paras.Count(n => ((Paragraph)n).ListFormat.ListLevelNumber == 2), Is.EqualTo(3));
        }

        [Test]
        public void ApplyNewListToParagraphs()
        {
            //ExStart
            //ExFor:ListCollection.Add(ListTemplate)
            //ExSummary:Shows how to create a list by applying a new list format to a collection of paragraphs.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Paragraph 1");
            builder.Writeln("Paragraph 2");
            builder.Write("Paragraph 3");

            NodeCollection paras = doc.GetChildNodes(NodeType.Paragraph, true);

            Assert.That(paras.Count(n => ((Paragraph)n).ListFormat.IsListItem), Is.EqualTo(0));

            List docList = doc.Lists.Add(ListTemplate.NumberUppercaseLetterDot);

            foreach (Paragraph paragraph in paras.OfType<Paragraph>())
            {
                paragraph.ListFormat.List = docList;
                paragraph.ListFormat.ListLevelNumber = 1;
            }

            Assert.That(paras.Count(n => ((Paragraph)n).ListFormat.IsListItem), Is.EqualTo(3));
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            paras = doc.GetChildNodes(NodeType.Paragraph, true);

            Assert.That(paras.Count(n => ((Paragraph)n).ListFormat.IsListItem), Is.EqualTo(3));
            Assert.That(paras.Count(n => ((Paragraph)n).ListFormat.ListLevelNumber == 1), Is.EqualTo(3));
        }

        //ExStart
        //ExFor:ListTemplate
        //ExSummary:Shows how to create a document that contains all outline headings list templates.
        [Test] //ExSkip
        public void OutlineHeadingTemplates()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            List docList = doc.Lists.Add(ListTemplate.OutlineHeadingsArticleSection);
            AddOutlineHeadingParagraphs(builder, docList, "Aspose.Words Outline - \"Article Section\"");

            docList = doc.Lists.Add(ListTemplate.OutlineHeadingsLegal);
            AddOutlineHeadingParagraphs(builder, docList, "Aspose.Words Outline - \"Legal\"");

            builder.InsertBreak(BreakType.PageBreak);

            docList = doc.Lists.Add(ListTemplate.OutlineHeadingsNumbers);
            AddOutlineHeadingParagraphs(builder, docList, "Aspose.Words Outline - \"Numbers\"");

            docList = doc.Lists.Add(ListTemplate.OutlineHeadingsChapter);
            AddOutlineHeadingParagraphs(builder, docList, "Aspose.Words Outline - \"Chapters\"");

            doc.Save(ArtifactsDir + "Lists.OutlineHeadingTemplates.docx");
            TestOutlineHeadingTemplates(new Document(ArtifactsDir + "Lists.OutlineHeadingTemplates.docx")); //ExSkip
        }

        private static void AddOutlineHeadingParagraphs(DocumentBuilder builder, List docList, string title)
        {
            builder.ParagraphFormat.ClearFormatting();
            builder.Writeln(title);

            for (int i = 0; i < 9; i++)
            {
                builder.ListFormat.List = docList;
                builder.ListFormat.ListLevelNumber = i;

                string styleName = "Heading " + (i + 1);
                builder.ParagraphFormat.StyleName = styleName;
                builder.Writeln(styleName);
            }

            builder.ListFormat.RemoveNumbers();
        }
        //ExEnd

        private void TestOutlineHeadingTemplates(Document doc)
        {
            List docList = doc.Lists[0]; // Article section list template.

            TestUtil.VerifyListLevel("Article \0.", 0.0d, NumberStyle.UppercaseRoman, docList.ListLevels[0]);
            TestUtil.VerifyListLevel("Section \0.\u0001", 0.0d, NumberStyle.LeadingZero, docList.ListLevels[1]);
            TestUtil.VerifyListLevel("(\u0002)", 14.4d, NumberStyle.LowercaseLetter, docList.ListLevels[2]);
            TestUtil.VerifyListLevel("(\u0003)", 36.0d, NumberStyle.LowercaseRoman, docList.ListLevels[3]);
            TestUtil.VerifyListLevel("\u0004)", 28.8d, NumberStyle.Arabic, docList.ListLevels[4]);
            TestUtil.VerifyListLevel("\u0005)", 36.0d, NumberStyle.LowercaseLetter, docList.ListLevels[5]);
            TestUtil.VerifyListLevel("\u0006)", 50.4d, NumberStyle.LowercaseRoman, docList.ListLevels[6]);
            TestUtil.VerifyListLevel("\a.", 50.4d, NumberStyle.LowercaseLetter, docList.ListLevels[7]);
            TestUtil.VerifyListLevel("\b.", 72.0d, NumberStyle.LowercaseRoman, docList.ListLevels[8]);

            docList = doc.Lists[1]; // Legal list template.

            TestUtil.VerifyListLevel("\0", 0.0d, NumberStyle.Arabic, docList.ListLevels[0]);
            TestUtil.VerifyListLevel("\0.\u0001", 0.0d, NumberStyle.Arabic, docList.ListLevels[1]);
            TestUtil.VerifyListLevel("\0.\u0001.\u0002", 0.0d, NumberStyle.Arabic, docList.ListLevels[2]);
            TestUtil.VerifyListLevel("\0.\u0001.\u0002.\u0003", 0.0d, NumberStyle.Arabic, docList.ListLevels[3]);
            TestUtil.VerifyListLevel("\0.\u0001.\u0002.\u0003.\u0004", 0.0d, NumberStyle.Arabic, docList.ListLevels[4]);
            TestUtil.VerifyListLevel("\0.\u0001.\u0002.\u0003.\u0004.\u0005", 0.0d, NumberStyle.Arabic, docList.ListLevels[5]);
            TestUtil.VerifyListLevel("\0.\u0001.\u0002.\u0003.\u0004.\u0005.\u0006", 0.0d, NumberStyle.Arabic, docList.ListLevels[6]);
            TestUtil.VerifyListLevel("\0.\u0001.\u0002.\u0003.\u0004.\u0005.\u0006.\a", 0.0d, NumberStyle.Arabic, docList.ListLevels[7]);
            TestUtil.VerifyListLevel("\0.\u0001.\u0002.\u0003.\u0004.\u0005.\u0006.\a.\b", 0.0d, NumberStyle.Arabic, docList.ListLevels[8]);

            docList = doc.Lists[2]; // Numbered list template.

            TestUtil.VerifyListLevel("\0.", 0.0d, NumberStyle.UppercaseRoman, docList.ListLevels[0]);
            TestUtil.VerifyListLevel("\u0001.", 36.0d, NumberStyle.UppercaseLetter, docList.ListLevels[1]);
            TestUtil.VerifyListLevel("\u0002.", 72.0d, NumberStyle.Arabic, docList.ListLevels[2]);
            TestUtil.VerifyListLevel("\u0003)", 108.0d, NumberStyle.LowercaseLetter, docList.ListLevels[3]);
            TestUtil.VerifyListLevel("(\u0004)", 144.0d, NumberStyle.Arabic, docList.ListLevels[4]);
            TestUtil.VerifyListLevel("(\u0005)", 180.0d, NumberStyle.LowercaseLetter, docList.ListLevels[5]);
            TestUtil.VerifyListLevel("(\u0006)", 216.0d, NumberStyle.LowercaseRoman, docList.ListLevels[6]);
            TestUtil.VerifyListLevel("(\a)", 252.0d, NumberStyle.LowercaseLetter, docList.ListLevels[7]);
            TestUtil.VerifyListLevel("(\b)", 288.0d, NumberStyle.LowercaseRoman, docList.ListLevels[8]);

            docList = doc.Lists[3]; // Chapter list template.

            TestUtil.VerifyListLevel("Chapter \0", 0.0d, NumberStyle.Arabic, docList.ListLevels[0]);
            TestUtil.VerifyListLevel("", 0.0d, NumberStyle.None, docList.ListLevels[1]);
            TestUtil.VerifyListLevel("", 0.0d, NumberStyle.None, docList.ListLevels[2]);
            TestUtil.VerifyListLevel("", 0.0d, NumberStyle.None, docList.ListLevels[3]);
            TestUtil.VerifyListLevel("", 0.0d, NumberStyle.None, docList.ListLevels[4]);
            TestUtil.VerifyListLevel("", 0.0d, NumberStyle.None, docList.ListLevels[5]);
            TestUtil.VerifyListLevel("", 0.0d, NumberStyle.None, docList.ListLevels[6]);
            TestUtil.VerifyListLevel("", 0.0d, NumberStyle.None, docList.ListLevels[7]);
            TestUtil.VerifyListLevel("", 0.0d, NumberStyle.None, docList.ListLevels[8]);
        }

        //ExStart
        //ExFor:ListCollection
        //ExFor:ListCollection.AddCopy(List)
        //ExSummary:Shows how to create a document with a sample of all the lists from another document.
        [Test] //ExSkip
        public void PrintOutAllLists()
        {
            Document srcDoc = new Document(MyDir + "Rendering.docx");

            Document dstDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(dstDoc);

            foreach (List srcList in srcDoc.Lists)
            {
                List dstList = dstDoc.Lists.AddCopy(srcList);
                AddListSample(builder, dstList);
            }

            dstDoc.Save(ArtifactsDir + "Lists.PrintOutAllLists.docx");
            TestPrintOutAllLists(srcDoc, new Document(ArtifactsDir + "Lists.PrintOutAllLists.docx")); //ExSkip
        }

        private static void AddListSample(DocumentBuilder builder, List docList)
        {
            builder.Writeln("Sample formatting of list with ListId:" + docList.ListId);
            builder.ListFormat.List = docList;
            for (int i = 0; i < docList.ListLevels.Count; i++)
            {
                builder.ListFormat.ListLevelNumber = i;
                builder.Writeln("Level " + i);
            }

            builder.ListFormat.RemoveNumbers();
            builder.Writeln();
        }
        //ExEnd

        private void TestPrintOutAllLists(Document listSourceDoc, Document outDoc)
        {
            foreach (List docList in outDoc.Lists)
                for (int i = 0; i < docList.ListLevels.Count; i++)
                {
                    ListLevel expectedListLevel = listSourceDoc.Lists.First(l => l.ListId == docList.ListId).ListLevels[i];
                    Assert.That(docList.ListLevels[i].NumberFormat, Is.EqualTo(expectedListLevel.NumberFormat));
                    Assert.That(docList.ListLevels[i].NumberPosition, Is.EqualTo(expectedListLevel.NumberPosition));
                    Assert.That(docList.ListLevels[i].NumberStyle, Is.EqualTo(expectedListLevel.NumberStyle));
                }
        }

        [Test]
        public void ListDocument()
        {
            //ExStart
            //ExFor:ListCollection.Document
            //ExFor:ListCollection.Count
            //ExFor:ListCollection.Item(Int32)
            //ExFor:ListCollection.GetListByListId
            //ExFor:List.Document
            //ExFor:List.ListId
            //ExSummary:Shows how to verify owner document properties of lists.
            Document doc = new Document();

            ListCollection lists = doc.Lists;
            Assert.That(lists.Document, Is.EqualTo(doc));

            List docList = lists.Add(ListTemplate.BulletDefault);
            Assert.That(docList.Document, Is.EqualTo(doc));

            Console.WriteLine("Current list count: " + lists.Count);
            Console.WriteLine("Is the first document list: " + (lists[0].Equals(docList)));
            Console.WriteLine("ListId: " + docList.ListId);
            Console.WriteLine("List is the same by ListId: " + (lists.GetListByListId(1).Equals(docList)));
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            lists = doc.Lists;
            
            Assert.That(lists.Document, Is.EqualTo(doc));
            Assert.That(lists.Count, Is.EqualTo(1));
            Assert.That(lists[0].ListId, Is.EqualTo(1));
            Assert.That(lists.GetListByListId(1), Is.EqualTo(lists[0]));
        }
        
        [Test]
        public void CreateListRestartAfterHigher()
        {
            //ExStart
            //ExFor:ListLevel.NumberStyle
            //ExFor:ListLevel.NumberFormat
            //ExFor:ListLevel.IsLegal
            //ExFor:ListLevel.RestartAfterLevel
            //ExFor:ListLevel.LinkedStyle
            //ExSummary:Shows advances ways of customizing list labels.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
            // We can create nested lists by increasing the indent level. 
            // We can begin and end a list by using a document builder's "ListFormat" property. 
            // Each paragraph that we add between a list's start and the end will become an item in the list.
            List docList = doc.Lists.Add(ListTemplate.NumberDefault);

            // Level 1 labels will be formatted according to the "Heading 1" paragraph style and will have a prefix.
            // These will look like "Appendix A", "Appendix B"...
            docList.ListLevels[0].NumberFormat = "Appendix \x0000";
            docList.ListLevels[0].NumberStyle = NumberStyle.UppercaseLetter;
            docList.ListLevels[0].LinkedStyle = doc.Styles["Heading 1"];

            // Level 2 labels will display the current numbers of the first and the second list levels and have leading zeroes.
            // If the first list level is at 1, then the list labels from these will look like "Section (1.01)", "Section (1.02)"...
            docList.ListLevels[1].NumberFormat = "Section (\x0000.\x0001)";
            docList.ListLevels[1].NumberStyle = NumberStyle.LeadingZero;

            // Note that the higher-level uses UppercaseLetter numbering.
            // We can set the "IsLegal" property to use Arabic numbers for the higher list levels.
            docList.ListLevels[1].IsLegal = true;
            docList.ListLevels[1].RestartAfterLevel = 0;

            // Level 3 labels will be upper case Roman numerals with a prefix and a suffix and will restart at each List level 1 item.
            // These list labels will look like "-I-", "-II-"...
            docList.ListLevels[2].NumberFormat = "-\x0002-";
            docList.ListLevels[2].NumberStyle = NumberStyle.UppercaseRoman;
            docList.ListLevels[2].RestartAfterLevel = 1;

            // Make labels of all list levels bold.
            foreach (ListLevel level in docList.ListLevels)
                level.Font.Bold = true;

            // Apply list formatting to the current paragraph.
            builder.ListFormat.List = docList;

            // Create list items that will display all three of our list levels.
            for (int n = 0; n < 2; n++)
            {
                for (int i = 0; i < 3; i++)
                {
                    builder.ListFormat.ListLevelNumber = i;
                    builder.Writeln("Level " + i);
                }
            }

            builder.ListFormat.RemoveNumbers();

            doc.Save(ArtifactsDir + "Lists.CreateListRestartAfterHigher.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Lists.CreateListRestartAfterHigher.docx");

            ListLevel listLevel = doc.Lists[0].ListLevels[0];

            TestUtil.VerifyListLevel("Appendix \0", 18.0d, NumberStyle.UppercaseLetter, listLevel);
            Assert.That(listLevel.IsLegal, Is.False);
            Assert.That(listLevel.RestartAfterLevel, Is.EqualTo(-1));
            Assert.That(listLevel.LinkedStyle.Name, Is.EqualTo("Heading 1"));

            listLevel = doc.Lists[0].ListLevels[1];

            TestUtil.VerifyListLevel("Section (\0.\u0001)", 54.0d, NumberStyle.LeadingZero, listLevel);
            Assert.That(listLevel.IsLegal, Is.True);
            Assert.That(listLevel.RestartAfterLevel, Is.EqualTo(0));
            Assert.That(listLevel.LinkedStyle, Is.Null);
        }

        [Test]
        public void GetListLabels()
        {
            //ExStart
            //ExFor:Document.UpdateListLabels()
            //ExFor:Node.ToString(SaveFormat)
            //ExFor:ListLabel
            //ExFor:Paragraph.ListLabel
            //ExFor:ListLabel.LabelValue
            //ExFor:ListLabel.LabelString
            //ExSummary:Shows how to extract the list labels of all paragraphs that are list items.
            Document doc = new Document(MyDir + "Rendering.docx");
            doc.UpdateListLabels();

            NodeCollection paras = doc.GetChildNodes(NodeType.Paragraph, true);

            // Find if we have the paragraph list. In our document, our list uses plain Arabic numbers,
            // which start at three and ends at six.
            foreach (Paragraph paragraph in paras.OfType<Paragraph>().Where(p => p.ListFormat.IsListItem).ToList())
            {
                Console.WriteLine($"List item paragraph #{paras.IndexOf(paragraph)}");

                // This is the text we get when getting when we output this node to text format.
                // This text output will omit list labels. Trim any paragraph formatting characters. 
                string paragraphText = paragraph.ToString(SaveFormat.Text).Trim();
                Console.WriteLine($"\tExported Text: {paragraphText}");

                ListLabel label = paragraph.ListLabel;

                // This gets the position of the paragraph in the current level of the list. If we have a list with multiple levels,
                // this will tell us what position it is on that level.
                Console.WriteLine($"\tNumerical Id: {label.LabelValue}");

                // Combine them together to include the list label with the text in the output.
                Console.WriteLine($"\tList label combined with text: {label.LabelString} {paragraphText}");
            }
            //ExEnd

            Assert.That(paras.OfType<Paragraph>().Count(p => p.ListFormat.IsListItem), Is.EqualTo(10));
        }

        [Test, Category("IgnoreOnJenkins")]
        public void CreatePictureBullet()
        {
            //ExStart
            //ExFor:ListLevel.CreatePictureBullet
            //ExFor:ListLevel.DeletePictureBullet
            //ExSummary:Shows how to set a custom image icon for list item labels.
            Document doc = new Document();

            List docList = doc.Lists.Add(ListTemplate.BulletCircle);

            // Create a picture bullet for the current list level, and set an image from a local file system
            // as the icon that the bullets for this list level will display.
            docList.ListLevels[0].CreatePictureBullet();
            docList.ListLevels[0].ImageData.SetImage(ImageDir + "Logo icon.ico");

            Assert.That(docList.ListLevels[0].ImageData.HasImage, Is.True);

            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ListFormat.List = docList;
            builder.Writeln("Hello world!");
            builder.Write("Hello again!");

            doc.Save(ArtifactsDir + "Lists.CreatePictureBullet.docx");

            docList.ListLevels[0].DeletePictureBullet();

            Assert.That(docList.ListLevels[0].ImageData, Is.Null);
            //ExEnd

            doc = new Document(ArtifactsDir + "Lists.CreatePictureBullet.docx");

            Assert.That(doc.Lists[0].ListLevels[0].ImageData.HasImage, Is.True);
        }

        [Test]
        public void GetCustomNumberStyleFormat()
        {
            //ExStart
            //ExFor:ListLevel.CustomNumberStyleFormat
            //ExFor:ListLevel.GetEffectiveValue(Int32, NumberStyle, String)
            //ExSummary:Shows how to get the format for a list with the custom number style.
            Document doc = new Document(MyDir + "List with leading zero.docx");

            ListLevel listLevel = doc.FirstSection.Body.Paragraphs[0].ListFormat.ListLevel;

            string customNumberStyleFormat = string.Empty;

            if (listLevel.NumberStyle == NumberStyle.Custom)
                customNumberStyleFormat = listLevel.CustomNumberStyleFormat;

            Assert.That(customNumberStyleFormat, Is.EqualTo("001, 002, 003, ..."));

            // We can get value for the specified index of the list item.
            Assert.That(ListLevel.GetEffectiveValue(4, NumberStyle.LowercaseRoman, null), Is.EqualTo("iv"));
            Assert.That(ListLevel.GetEffectiveValue(5, NumberStyle.Custom, customNumberStyleFormat), Is.EqualTo("005"));
            //ExEnd

            Assert.Throws<ArgumentException>(
                () => ListLevel.GetEffectiveValue(5, NumberStyle.LowercaseRoman, customNumberStyleFormat));
            Assert.Throws<ArgumentException>(() => ListLevel.GetEffectiveValue(5, NumberStyle.Custom, null));
            Assert.Throws<ArgumentException>(() => ListLevel.GetEffectiveValue(5, NumberStyle.Custom, "...."));
        }

        [Test]
        public void HasSameTemplate()
        {
            //ExStart
            //ExFor:List.HasSameTemplate(List)
            //ExSummary:Shows how to define lists with the same ListDefId.
            Document doc = new Document(MyDir + "Different lists.docx");

            Assert.That(doc.Lists[0].HasSameTemplate(doc.Lists[1]), Is.True);
            Assert.That(doc.Lists[1].HasSameTemplate(doc.Lists[2]), Is.False);
            //ExEnd
        }

        [Test]
        public void SetCustomNumberStyleFormat()
        {
            //ExStart:SetCustomNumberStyleFormat
            //GistId:ac8ba4eb35f3fbb8066b48c999da63b0
            //ExFor:ListLevel.CustomNumberStyleFormat
            //ExSummary:Shows how to set customer number style format.
            Document doc = new Document(MyDir + "List with leading zero.docx");

            doc.UpdateListLabels();

            ParagraphCollection paras = doc.FirstSection.Body.Paragraphs;
            Assert.That(paras[0].ListLabel.LabelString, Is.EqualTo("001."));
            Assert.That(paras[1].ListLabel.LabelString, Is.EqualTo("0001."));
            Assert.That(paras[2].ListLabel.LabelString, Is.EqualTo("0002."));

            paras[1].ListFormat.ListLevel.CustomNumberStyleFormat = "001, 002, 003, ...";

            doc.UpdateListLabels();

            Assert.That(paras[0].ListLabel.LabelString, Is.EqualTo("001."));
            Assert.That(paras[1].ListLabel.LabelString, Is.EqualTo("001."));
            Assert.That(paras[2].ListLabel.LabelString, Is.EqualTo("002."));
            //ExEnd:SetCustomNumberStyleFormat
        }

        [Test]
        public void AddSingleLevelList()
        {
            //ExStart:AddSingleLevelList
            //GistId:95fdae949cefbf2ce485acc95cccc495
            //ExFor:ListCollection.AddSingleLevelList(ListTemplate)
            //ExSummary:Shows how to create a new single level list based on the predefined template.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            ListCollection listCollection = doc.Lists;

            // Creates the bulleted list from BulletCircle template.
            List bulletedList = listCollection.AddSingleLevelList(ListTemplate.BulletCircle);

            // Writes the bulleted list to the resulting document.
            builder.Writeln("Bulleted list starts below:");
            builder.ListFormat.List = bulletedList;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();

            // Creates the numbered list from NumberUppercaseLetterDot template.
            List numberedList = listCollection.AddSingleLevelList(ListTemplate.NumberUppercaseLetterDot);

            // Writes the numbered list to the resulting document.
            builder.Writeln("Numbered list starts below:");
            builder.ListFormat.List = numberedList;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");

            doc.Save(ArtifactsDir + "Lists.AddSingleLevelList.docx");
            //ExEnd:AddSingleLevelList
        }
    }
}