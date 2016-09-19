// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Drawing;

using Aspose.Words;
using Aspose.Words.Lists;

using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExLists : ApiExampleBase
    {
        private readonly string _image = MyDir + "Test_636_852.gif";

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
            //ExSummary:Shows how to apply default bulleted or numbered list formatting to paragraphs when using DocumentBuilder.

            DocumentBuilder builder = new DocumentBuilder();

            builder.Writeln("Aspose.Words allows:");
            builder.Writeln();

            // Start a numbered list with default formatting.
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("Opening documents from different formats:");

            // Go to second list level, add more text.
            builder.ListFormat.ListIndent();
            builder.Writeln("DOC");
            builder.Writeln("PDF");
            builder.Writeln("HTML");

            // Outdent to the first list level.
            builder.ListFormat.ListOutdent();
            builder.Writeln("Processing documents");
            builder.Writeln("Saving documents in different formats:");

            // Indent the list level again.
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
            builder.Writeln();

            builder.Writeln("Aspose.Words main advantages are:");
            builder.Writeln();

            // Start a bulleted list with default formatting.
            builder.ListFormat.ApplyBulletDefault();
            builder.Writeln("Great performance");
            builder.Writeln("High reliability");
            builder.Writeln("Quality code and working");
            builder.Writeln("Wide variety of features");
            builder.Writeln("Easy to understand API");

            // End the bulleted list.
            builder.ListFormat.RemoveNumbers();

            builder.Document.Save(MyDir + @"\Artifacts\Lists.ApplyDefaultBulletsAndNumbers.doc");
            //ExEnd
        }

        [Test]
        public void SpecifyListLevel()
        {
            //ExStart
            //ExFor:ListCollection
            //ExFor:List
            //ExFor:ListFormat
            //ExFor:ListFormat.ListLevelNumber
            //ExFor:ListFormat.List
            //ExFor:ListTemplate
            //ExFor:DocumentBase.Lists
            //ExFor:ListCollection.Add(ListTemplate)
            //ExSummary:Shows how to specify list level number when building a list using DocumentBuilder.
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

            builder.Document.Save(MyDir + @"\Artifacts\Lists.SpecifyListLevel.doc");
            //ExEnd
        }

        [Test]
        public void NestedLists()
        {
            //ExStart
            //ExFor:ListFormat.List
            //ExSummary:Shows how to start a numbered list, add a bulleted list inside it, then return to the numbered list.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create an outline list for the headings.
            Aspose.Words.Lists.List outlineList = doc.Lists.Add(ListTemplate.OutlineNumbers);
            builder.ListFormat.List = outlineList;
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("This is my Chapter 1");

            // Create a numbered list.
            Aspose.Words.Lists.List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);
            builder.ListFormat.List = numberedList;
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Numebered list item 1.");

            // Create a bulleted list.
            Aspose.Words.Lists.List bulletedList = doc.Lists.Add(ListTemplate.BulletDefault);
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

            builder.Document.Save(MyDir + @"\Artifacts\Lists.NestedLists.doc");
            //ExEnd
        }

        [Test]
        public void CreateCustomList()
        {
            //ExStart
            //ExFor:List
            //ExFor:List.ListLevels
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

            // Create a list based on one of the Microsoft Word list templates.
            Aspose.Words.Lists.List list = doc.Lists.Add(ListTemplate.NumberDefault);

            // Completely customize one list level.
            ListLevel level1 = list.ListLevels[0];
            level1.Font.Color = Color.Red;
            level1.Font.Size = 24;
            level1.NumberStyle = NumberStyle.OrdinalText;
            level1.StartAt = 21;
            level1.NumberFormat = "\x0000";

            level1.NumberPosition = -36;
            level1.TextPosition = 144;
            level1.TabPosition = 144;

            // Completely customize yet another list level.
            ListLevel level2 = list.ListLevels[1];
            level2.Alignment = ListLevelAlignment.Right;
            level2.NumberStyle = NumberStyle.Bullet;
            level2.Font.Name = "Wingdings";
            level2.Font.Color = Color.Blue;
            level2.Font.Size = 24;
            level2.NumberFormat = "\xf0af";	// A bullet that looks like some sort of a star.
            level2.TrailingCharacter = ListTrailingCharacter.Space;
            level2.NumberPosition = 144;

            // Now add some text that uses the list that we created.			
            // It does not matter when to customize the list - before or after adding the paragraphs.
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ListFormat.List = list;
            builder.Writeln("The quick brown fox...");
            builder.Writeln("The quick brown fox...");

            builder.ListFormat.ListIndent();
            builder.Writeln("jumped over the lazy dog.");
            builder.Writeln("jumped over the lazy dog.");

            builder.ListFormat.ListOutdent();
            builder.Writeln("The quick brown fox...");

            builder.ListFormat.RemoveNumbers();

            builder.Document.Save(MyDir + @"\Artifacts\Lists.CreateCustomList.doc");
            //ExEnd
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
            //ExFor:ListFormat.List
            //ExSummary:Shows how to restart numbering in a list by copying a list.
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

            builder.Document.Save(MyDir + @"\Artifacts\Lists.RestartNumberingUsingListCopy.doc");
            //ExEnd
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

            // Create a new list style. 
            // List formatting associated with this list style is default numbered.
            Style listStyle = doc.Styles.Add(StyleType.List, "MyListStyle");

            // This list defines the formatting of the list style.
            // Note this list can not be used directly to apply formatting to paragraphs (see below).
            Aspose.Words.Lists.List list1 = listStyle.List;

            // Check some basic rules about the list that defines a list style.
            Console.WriteLine("IsListStyleDefinition: " + list1.IsListStyleDefinition);
            Console.WriteLine("IsListStyleReference: " + list1.IsListStyleReference);
            Console.WriteLine("IsMultiLevel: " + list1.IsMultiLevel);
            Console.WriteLine("List style has been set: " + (listStyle == list1.Style));

            // Modify formatting of the list style to our liking.
            for (int i = 0; i < list1.ListLevels.Count; i++)
            {
                ListLevel level = list1.ListLevels[i];
                level.Font.Name = "Verdana";
                level.Font.Color = Color.Blue;
                level.Font.Bold = true;
            }


            // Add some text to our document and use the list style.
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Using list style first time:");

            // This creates a list based on the list style.
            Aspose.Words.Lists.List list2 = doc.Lists.Add(listStyle);

            // Check some basic rules about the list that references a list style.
            Console.WriteLine("IsListStyleDefinition: " + list2.IsListStyleDefinition);
            Console.WriteLine("IsListStyleReference: " + list2.IsListStyleReference);
            Console.WriteLine("List Style has been set: " + (listStyle == list2.Style));

            // Apply the list that references the list style.
            builder.ListFormat.List = list2;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();


            builder.Writeln("Using list style second time:");

            // Create and apply another list based on the list style.
            Aspose.Words.Lists.List list3 = doc.Lists.Add(listStyle);
            builder.ListFormat.List = list3;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();

            builder.Document.Save(MyDir + @"\Artifacts\Lists.CreateAndUseListStyle.doc");
            //ExEnd

            // Verify properties of list 1
            Assert.IsTrue(list1.IsListStyleDefinition);
            Assert.IsFalse(list1.IsListStyleReference);
            Assert.IsTrue(list1.IsMultiLevel);
            Assert.AreEqual(listStyle, list1.Style);

            // Verify properties of list 2
            Assert.IsFalse(list2.IsListStyleDefinition);
            Assert.IsTrue(list2.IsListStyleReference);
            Assert.AreEqual(listStyle, list2.Style);
        }

        [Test]
        public void DetectBulletedParagraphs()
        {
            Document doc = new Document();

            //ExStart
            //ExFor:Paragraph.ListFormat
            //ExFor:ListFormat.IsListItem
            //ExFor:CompositeNode.GetText
            //ExFor:List.ListId
            //ExSummary:Finds and outputs all paragraphs in a document that are bulleted or numbered.
            NodeCollection paras = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph para in paras)
            {
                if (para.ListFormat.IsListItem)
                {
                    Console.WriteLine(string.Format("*** A paragraph belongs to list {0}", para.ListFormat.List.ListId));
                    Console.WriteLine(para.GetText());
                }
            }
            //ExEnd
        }

        [Test]
        public void RemoveBulletsFromParagraphs()
        {
            Document doc = new Document();

            //ExStart
            //ExFor:Paragraph.ListFormat
            //ExFor:ListFormat.RemoveNumbers
            //ExSummary:Removes bullets and numbering from all paragraphs in the main text of a section.
            Body body = doc.FirstSection.Body;

            foreach (Paragraph paragraph in body.Paragraphs)
                paragraph.ListFormat.RemoveNumbers();

            //ExEnd
        }

        [Test]
        public void ApplyExistingListToParagraphs()
        {
            Document doc = new Document();
            doc.Lists.Add(ListTemplate.NumberDefault);

            //ExStart
            //ExFor:Paragraph.ListFormat
            //ExFor:ListFormat.List
            //ExFor:ListFormat.ListLevelNumber
            //ExFor:ListCollection.Item(Int32)
            //ExSummary:Applies list formatting of an existing list to a collection of paragraphs.
            Body body = doc.FirstSection.Body;
            Aspose.Words.Lists.List list = doc.Lists[0];
            foreach (Paragraph paragraph in body.Paragraphs)
            {
                paragraph.ListFormat.List = list;
                paragraph.ListFormat.ListLevelNumber = 2;
            }
            //ExEnd
        }

        [Test]
        public void ApplyNewListToParagraphs()
        {
            Document doc = new Document();

            //ExStart
            //ExFor:Paragraph.ListFormat
            //ExFor:ListFormat.ListLevelNumber
            //ExFor:ListCollection.Add(ListTemplate)
            //ExSummary:Creates new list formatting and applies it to a collection of paragraphs.
            Aspose.Words.Lists.List list = doc.Lists.Add(ListTemplate.NumberUppercaseLetterDot);

            Body body = doc.FirstSection.Body;
            foreach (Paragraph paragraph in body.Paragraphs)
            {
                paragraph.ListFormat.List = list;
                paragraph.ListFormat.ListLevelNumber = 1;
            }
            //ExEnd
        }

        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void OutlineHeadingTemplatesCaller()
        {
            this.OutlineHeadingTemplates();
        }
        
        //ExStart
        //ExFor:ListTemplate
        //ExSummary:Creates a sample document that exercises all outline headings list templates.
        public void OutlineHeadingTemplates()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Aspose.Words.Lists.List list = doc.Lists.Add(ListTemplate.OutlineHeadingsArticleSection);
            AddOutlineHeadingParagraphs(builder, list, "Aspose.Words Outline 1");

            list = doc.Lists.Add(ListTemplate.OutlineHeadingsLegal);
            AddOutlineHeadingParagraphs(builder, list, "Aspose.Words Outline 2");

            builder.InsertBreak(BreakType.PageBreak);

            list = doc.Lists.Add(ListTemplate.OutlineHeadingsNumbers);
            AddOutlineHeadingParagraphs(builder, list, "Aspose.Words Outline 3");

            list = doc.Lists.Add(ListTemplate.OutlineHeadingsChapter);
            AddOutlineHeadingParagraphs(builder, list, "Aspose.Words Outline 4");

            builder.Document.Save(MyDir + @"\Artifacts\Lists.OutlineHeadingTemplates.doc");
        }

        private static void AddOutlineHeadingParagraphs(DocumentBuilder builder, Aspose.Words.Lists.List list, string title)
        {
            builder.ParagraphFormat.ClearFormatting();
            builder.Writeln(title);

            for (int i = 0; i < 9; i++)
            {
                builder.ListFormat.List = list;
                builder.ListFormat.ListLevelNumber = i;

                string styleName = "Heading " + (i + 1).ToString();
                builder.ParagraphFormat.StyleName = styleName;
                builder.Writeln(styleName);
            }

            builder.ListFormat.RemoveNumbers();
        }
        //ExEnd

        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void PrintOutAllListsCaller()
        {
            this.PrintOutAllLists();
        }
        
        //ExStart
        //ExFor:ListCollection
        //ExFor:ListCollection.AddCopy(List)
        //ExFor:ListCollection.GetEnumerator
        //ExSummary:Enumerates through all lists defined in one document and creates a sample of those lists in another document.
        public void PrintOutAllLists()
        {
            // You can use any of your documents to try this little program out.
            Document srcDoc = new Document(MyDir + "Lists.PrintOutAllLists.doc");

            // This will be the sample document we product.
            Document dstDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(dstDoc);

            foreach (Aspose.Words.Lists.List srcList in srcDoc.Lists)
            {
                // This copies the list formatting from the source into the destination document.
                Aspose.Words.Lists.List dstList = dstDoc.Lists.AddCopy(srcList);
                AddListSample(builder, dstList);
            }

            dstDoc.Save(MyDir + @"\Artifacts\Lists.PrintOutAllLists.doc");
        }

        private static void AddListSample(DocumentBuilder builder, Aspose.Words.Lists.List list)
        {
            builder.Writeln("Sample formatting of list with ListId:" + list.ListId);
            builder.ListFormat.List = list;
            for (int i = 0; i < list.ListLevels.Count; i++)
            {
                builder.ListFormat.ListLevelNumber = i;
                builder.Writeln("Level " + i);
            }
            builder.ListFormat.RemoveNumbers();
            builder.Writeln();
        }
        //ExEnd		

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
            //ExSummary:Illustrates the owner document properties of lists.
            Document doc = new Document();

            ListCollection lists = doc.Lists;
            // All of these should be equal.
            Console.WriteLine("ListCollection document is doc: " + (doc == lists.Document));
            Console.WriteLine("Starting list count: " + lists.Count);

            Aspose.Words.Lists.List list = lists.Add(ListTemplate.BulletDefault);
            Console.WriteLine("List document is doc: " + (list.Document == doc));
            Console.WriteLine("List count after adding list: " + lists.Count);
            Console.WriteLine("Is the first document list: " + (lists[0] == list));
            Console.WriteLine("ListId: " + list.ListId);
            Console.WriteLine("List is the same by ListId: " + (lists.GetListByListId(1) == list));
            //ExEnd

            // Verify these properties
            Assert.AreEqual(doc, lists.Document);
            Assert.AreEqual(doc, list.Document);
            Assert.AreEqual(1, lists.Count);
            Assert.AreEqual(list, lists[0]);
            Assert.AreEqual(1, list.ListId);
            Assert.AreEqual(list, lists.GetListByListId(1));
        }

        [Test]
        public void ListFormatListLevel()
        {
            //ExStart
            //ExFor:ListFormat.ListLevel
            //ExSummary:Shows how to modify list formatting of the current list level.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create and apply list formatting to the current paragraph.
            builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

            // Modify formatting of the current (first) list level.
            builder.ListFormat.ListLevel.Font.Bold = true;

            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();
            //ExEnd
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
            //ExFor:ListLevelCollection.GetEnumerator
            //ExSummary:Shows how to create a list with some advanced formatting.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Aspose.Words.Lists.List list = doc.Lists.Add(ListTemplate.NumberDefault);

            // Level 1 labels will be "Appendix A", continuous and linked to the Heading 1 paragraph style.
            list.ListLevels[0].NumberFormat = "Appendix \x0000";
            list.ListLevels[0].NumberStyle = NumberStyle.UppercaseLetter;
            list.ListLevels[0].LinkedStyle = doc.Styles["Heading 1"];

            // Level 2 labels will be "Section (1.01)" and restarting after Level 2 item appears.
            list.ListLevels[1].NumberFormat = "Section (\x0000.\x0001)";
            list.ListLevels[1].NumberStyle = NumberStyle.LeadingZero;
            // Notice the higher level uses UppercaseLetter numbering, but we want arabic number
            // of the higher levels to appear in this level, therefore set this property.
            list.ListLevels[1].IsLegal = true;
            list.ListLevels[1].RestartAfterLevel = 0;

            // Level 3 labels will be "-I-" and restarting after Level 2 item appears.
            list.ListLevels[2].NumberFormat = "-\x0002-";
            list.ListLevels[2].NumberStyle = NumberStyle.UppercaseRoman;
            list.ListLevels[2].RestartAfterLevel = 1;

            // Make labels of all list levels bold.
            foreach (ListLevel level in list.ListLevels)
                level.Font.Bold = true;


            // Apply list formatting to the current paragraph.
            builder.ListFormat.List = list;

            // Exercise the 3 levels we created two times.
            for (int n = 0; n < 2; n++)
            {
                for (int i = 0; i < 3; i++)
                {
                    builder.ListFormat.ListLevelNumber = i;
                    builder.Writeln("Level " + i);
                }
            }

            builder.ListFormat.RemoveNumbers();

            builder.Document.Save(MyDir + @"\Artifacts\Lists.CreateListRestartAfterHigher.doc");
            //ExEnd
        }

        [Test]
        public void ParagraphStyleBulleted()
        {
            //ExStart
            //ExFor:StyleCollection
            //ExFor:DocumentBase.Styles
            //ExFor:Style
            //ExFor:Font
            //ExFor:Style.Font
            //ExFor:Style.ParagraphFormat
            //ExFor:Style.ListFormat
            //ExFor:ParagraphFormat.Style
            //ExSummary:Shows how to create and use a paragraph style with list formatting.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a paragraph style and specify some formatting for it.
            Style style = doc.Styles.Add(StyleType.Paragraph, "MyStyle1");
            style.Font.Size = 24;
            style.Font.Name = "Verdana";
            style.ParagraphFormat.SpaceAfter = 12;

            // Create a list and make sure the paragraphs that use this style will use this list.
            style.ListFormat.List = doc.Lists.Add(ListTemplate.BulletDefault);
            style.ListFormat.ListLevelNumber = 0;

            // Apply the paragraph style to the current paragraph in the document and add some text.
            builder.ParagraphFormat.Style = style;
            builder.Writeln("Hello World: MyStyle1, bulleted.");

            // Change to a paragraph style that has no list formatting.
            builder.ParagraphFormat.Style = doc.Styles["Normal"];
            builder.Writeln("Hello World: Normal.");

            builder.Document.Save(MyDir + @"\Artifacts\Lists.ParagraphStyleBulleted.doc");
            //ExEnd
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
            //ExSummary:Shows how to extract the label of each paragraph in a list as a value or a string.
            Document doc = new Document(MyDir + "Lists.PrintOutAllLists.doc");
            doc.UpdateListLabels();
            int listParaCount = 1;

            foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                // Find if we have the paragraph list. In our document our list uses plain arabic numbers,
                // which start at three and ends at six.
                if (paragraph.ListFormat.IsListItem)
                {
                    Console.WriteLine("Paragraph #{0}", listParaCount);

                    // This is the text we get when actually getting when we output this node to text format. 
                    // The list labels are not included in this text output. Trim any paragraph formatting characters.
                    string paragraphText = paragraph.ToString(SaveFormat.Text).Trim();
                    Console.WriteLine("Exported Text: " + paragraphText);

                    ListLabel label = paragraph.ListLabel;
                    // This gets the position of the paragraph in current level of the list. If we have a list with multiple level then this
                    // will tell us what position it is on that particular level.
                    Console.WriteLine("Numerical Id: " + label.LabelValue);

                    // Combine them together to include the list label with the text in the output.
                    Console.WriteLine("List label combined with text: " + label.LabelString + " " + paragraphText);

                    listParaCount++;
                }

            }
            //ExEnd
        }

        [Test]
        public void CreatePictureBullet()
        {
            //ExStart
            //ExFor: ListLevel.CreatePictureBullet
            //ExFor: ListLevel.DeletePictureBullet
            //ExSummary: Shows how to creating and deleting picture bullet with custom image
            Document doc = new Document();

            // Create a list with template
            Aspose.Words.Lists.List list = doc.Lists.Add(ListTemplate.BulletCircle);

            // Create picture bullet for the current list level
            list.ListLevels[0].CreatePictureBullet();

            // Set your own picture bullet image through the ImageData
            list.ListLevels[0].ImageData.SetImage(this._image);

            Assert.IsTrue(list.ListLevels[0].ImageData.HasImage);
            
            // Delete picture bullet
            list.ListLevels[0].DeletePictureBullet();
            
            Assert.IsNull(list.ListLevels[0].ImageData);
            //ExEnd
        }
    }
}
