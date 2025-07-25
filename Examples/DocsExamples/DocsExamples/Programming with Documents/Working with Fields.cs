﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents
{
    internal class WorkingWithFields : DocsExamplesBase
    {
        [Test]
        public void FieldCode()
        {
            //ExStart:FieldCode
            //GistId:7c2b7b650a88375b1d438746f78f0d64
            Document doc = new Document(MyDir + "Hyperlinks.docx");

            foreach (Field field in doc.Range.Fields)
            {
                string fieldCode = field.GetFieldCode();
                string fieldResult = field.Result;
            }
            //ExEnd:FieldCode
        }

        [Test]
        public void ChangeFieldUpdateCultureSource()
        {
            //ExStart:ChangeFieldUpdateCultureSource
            //GistId:9e90defe4a7bcafb004f73a2ef236986
            //ExStart:DocumentBuilderInsertField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert content with German locale.
            builder.Font.LocaleId = 1031;
            builder.InsertField("MERGEFIELD Date1 \\@ \"dddd, d MMMM yyyy\"");
            builder.Write(" - ");
            builder.InsertField("MERGEFIELD Date2 \\@ \"dddd, d MMMM yyyy\"");
            //ExEnd:DocumentBuilderInsertField

            // Shows how to specify where the culture used for date formatting during field update and mail merge is chosen from
            // set the culture used during field update to the culture used by the field.
            doc.FieldOptions.FieldUpdateCultureSource = FieldUpdateCultureSource.FieldCode;
            doc.MailMerge.Execute(new string[] { "Date2" }, new object[] { new DateTime(2011, 1, 01) });
            
            doc.Save(ArtifactsDir + "WorkingWithFields.ChangeFieldUpdateCultureSource.docx");
            //ExEnd:ChangeFieldUpdateCultureSource
        }

        [Test]
        public void SpecifyLocaleAtFieldLevel()
        {
            //ExStart:SpecifyLocaleAtFieldLevel
            //GistId:1cf07762df56f15067d6aef90b14b3db
            DocumentBuilder builder = new DocumentBuilder();

            Field field = builder.InsertField(FieldType.FieldDate, true);
            field.LocaleId = 1049;
            
            builder.Document.Save(ArtifactsDir + "WorkingWithFields.SpecifyLocaleAtFieldLevel.docx");
            //ExEnd:SpecifyLocaleAtFieldLevel
        }

        [Test]
        public void ReplaceHyperlinks()
        {
            //ExStart:ReplaceHyperlinks
            //GistId:0213851d47551e83af42233f4d075cf6
            Document doc = new Document(MyDir + "Hyperlinks.docx");

            foreach (Field field in doc.Range.Fields)
            {
                if (field.Type == FieldType.FieldHyperlink)
                {
                    FieldHyperlink hyperlink = (FieldHyperlink) field;

                    // Some hyperlinks can be local (links to bookmarks inside the document), ignore these.
                    if (hyperlink.SubAddress != null)
                        continue;

                    hyperlink.Address = "http://www.aspose.com";
                    hyperlink.Result = "Aspose - The .NET & Java Component Publisher";
                }
            }

            doc.Save(ArtifactsDir + "WorkingWithFields.ReplaceHyperlinks.docx");
            //ExEnd:ReplaceHyperlinks
        }

        [Test]
        public void RenameMergeFields()
        {
            //ExStart:RenameMergeFields
            //GistId:bf0f8a6b40b69a5274ab3553315e147f
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertField(@"MERGEFIELD MyMergeField1 \* MERGEFORMAT");
            builder.InsertField(@"MERGEFIELD MyMergeField2 \* MERGEFORMAT");

            foreach (Field f in doc.Range.Fields)
            {
                if (f.Type == FieldType.FieldMergeField)
                {
                    FieldMergeField mergeField = (FieldMergeField)f;
                    mergeField.FieldName = mergeField.FieldName + "_Renamed";
                    mergeField.Update();
                }
            }

            doc.Save(ArtifactsDir + "WorkingWithFields.RenameMergeFields.docx");
            //ExEnd:RenameMergeFields
        }

        [Test]
        public void RemoveField()
        {
            //ExStart:RemoveField
            //GistId:8c604665c1b97795df7a1e665f6b44ce
            Document doc = new Document(MyDir + "Various fields.docx");
            
            Field field = doc.Range.Fields[0];
            field.Remove();
            //ExEnd:RemoveField
        }

        [Test]
        public void UnlinkFields()
        {
            //ExStart:UnlinkFields
            //GistId:f3592014d179ecb43905e37b2a68bc92
            Document doc = new Document(MyDir + "Various fields.docx");
            doc.UnlinkFields();
            //ExEnd:UnlinkFields
        }

        [Test]
        public void InsertToaFieldWithoutDocumentBuilder()
        {
            //ExStart:InsertToaFieldWithoutDocumentBuilder
            //GistId:1cf07762df56f15067d6aef90b14b3db
            Document doc = new Document();
            Paragraph para = new Paragraph(doc);

            // We want to insert TA and TOA fields like this:
            // { TA  \c 1 \l "Value 0" }
            // { TOA  \c 1 }

            FieldTA fieldTA = (FieldTA) para.AppendField(FieldType.FieldTOAEntry, false);
            fieldTA.EntryCategory = "1";
            fieldTA.LongCitation = "Value 0";

            doc.FirstSection.Body.AppendChild(para);

            para = new Paragraph(doc);

            FieldToa fieldToa = (FieldToa) para.AppendField(FieldType.FieldTOA, false);
            fieldToa.EntryCategory = "1";
            doc.FirstSection.Body.AppendChild(para);

            fieldToa.Update();

            doc.Save(ArtifactsDir + "WorkingWithFields.InsertToaFieldWithoutDocumentBuilder.docx");
            //ExEnd:InsertToaFieldWithoutDocumentBuilder
        }

        [Test]
        public void InsertNestedFields()
        {
            //ExStart:InsertNestedFields
            //GistId:1cf07762df56f15067d6aef90b14b3db
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            for (int i = 0; i < 5; i++)
                builder.InsertBreak(BreakType.PageBreak);

            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

            // We want to insert a field like this:
            // { IF {PAGE} <> {NUMPAGES} "See Next Page" "Last Page" }
            Field field = builder.InsertField(@"IF ");
            builder.MoveTo(field.Separator);
            builder.InsertField("PAGE");
            builder.Write(" <> ");
            builder.InsertField("NUMPAGES");
            builder.Write(" \"See Next Page\" \"Last Page\" ");

            field.Update();

            doc.Save(ArtifactsDir + "WorkingWithFields.InsertNestedFields.docx");
            //ExEnd:InsertNestedFields
        }

        [Test]
        public void InsertMergeFieldUsingDom()
        {
            //ExStart:InsertMergeFieldUsingDom
            //GistId:1cf07762df56f15067d6aef90b14b3db
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Paragraph para = (Paragraph) doc.GetChild(NodeType.Paragraph, 0, true);
            builder.MoveTo(para);

            // We want to insert a merge field like this:
            // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v" }
            FieldMergeField field = (FieldMergeField) builder.InsertField(FieldType.FieldMergeField, false);
            // { " MERGEFIELD Test1" }
            field.FieldName = "Test1";
            // { " MERGEFIELD Test1 \\b Test2" }
            field.TextBefore = "Test2";
            // { " MERGEFIELD Test1 \\b Test2 \\f Test3 }
            field.TextAfter = "Test3";
            // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m" }
            field.IsMapped = true;
            // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v" }
            field.IsVerticalFormatting = true;

            field.Update();

            doc.Save(ArtifactsDir + "WorkingWithFields.InsertMergeFieldUsingDom.docx");
            //ExEnd:InsertMergeFieldUsingDom
        }

        [Test]
        public void InsertAddressBlockFieldUsingDom()
        {
            //ExStart:InsertAddressBlockFieldUsingDom
            //GistId:1cf07762df56f15067d6aef90b14b3db
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Paragraph para = (Paragraph) doc.GetChild(NodeType.Paragraph, 0, true);
            builder.MoveTo(para);

            // We want to insert a mail merge address block like this:
            // { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\" }
            FieldAddressBlock field = (FieldAddressBlock) builder.InsertField(FieldType.FieldAddressBlock, false);
            // { ADDRESSBLOCK \\c 1" }
            field.IncludeCountryOrRegionName = "1";
            // { ADDRESSBLOCK \\c 1 \\d" }
            field.FormatAddressOnCountryOrRegion = true;
            // { ADDRESSBLOCK \\c 1 \\d \\e Test2 }
            field.ExcludedCountryOrRegionName = "Test2";
            // { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 }
            field.NameAndAddressFormat = "Test3";
            // { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\" }
            field.LanguageId = "Test 4";

            field.Update();

            doc.Save(ArtifactsDir + "WorkingWithFields.InsertAddressBlockFieldUsingDom.docx");
            //ExEnd:InsertAddressBlockFieldUsingDom
        }

        [Test]
        public void InsertFieldIncludeTextWithoutDocumentBuilder()
        {
            //ExStart:InsertFieldIncludeTextWithoutDocumentBuilder
            //GistId:1cf07762df56f15067d6aef90b14b3db
            Document doc = new Document();

            Paragraph para = new Paragraph(doc);

            // We want to insert an INCLUDETEXT field like this:
            // { INCLUDETEXT  "file path" }
            FieldIncludeText fieldIncludeText = (FieldIncludeText) para.AppendField(FieldType.FieldIncludeText, false);
            fieldIncludeText.BookmarkName = "bookmark";
            fieldIncludeText.SourceFullName = MyDir + "IncludeText.docx";

            doc.FirstSection.Body.AppendChild(para);

            fieldIncludeText.Update();

            doc.Save(ArtifactsDir + "WorkingWithFields.InsertIncludeFieldWithoutDocumentBuilder.docx");
            //ExEnd:InsertFieldIncludeTextWithoutDocumentBuilder
        }

        [Test]
        public void InsertFieldNone()
        {
            //ExStart:InsertFieldNone
            //GistId:1cf07762df56f15067d6aef90b14b3db
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            FieldUnknown field = (FieldUnknown) builder.InsertField(FieldType.FieldNone, false);

            doc.Save(ArtifactsDir + "WorkingWithFields.InsertFieldNone.docx");
            //ExEnd:InsertFieldNone
        }

        [Test]
        public void InsertField()
        {
            //ExStart:InsertField
            //GistId:1cf07762df56f15067d6aef90b14b3db
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.InsertField(@"MERGEFIELD MyFieldName \* MERGEFORMAT");
            
            doc.Save(ArtifactsDir + "WorkingWithFields.InsertField.docx");
            //ExEnd:InsertField
        }

        [Test]
        public void InsertFieldUsingFieldBuilder()
        {
            //ExStart:InsertFieldUsingFieldBuilder
            //GistId:1cf07762df56f15067d6aef90b14b3db
            Document doc = new Document();

            // Prepare IF field with two nested MERGEFIELD fields: { IF "left expression" = "right expression" "Firstname: { MERGEFIELD firstname }" "Lastname: { MERGEFIELD lastname }"}
            FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FieldIf)
                .AddArgument("left expression")
                .AddArgument("=")
                .AddArgument("right expression")
                .AddArgument(
                    new FieldArgumentBuilder()
                        .AddText("Firstname: ")
                        .AddField(new FieldBuilder(FieldType.FieldMergeField).AddArgument("firstname")))
                .AddArgument(
                    new FieldArgumentBuilder()
                        .AddText("Lastname: ")
                        .AddField(new FieldBuilder(FieldType.FieldMergeField).AddArgument("lastname")));

            // Insert IF field in exact location
            Field field = fieldBuilder.BuildAndInsert(doc.FirstSection.Body.FirstParagraph);
            field.Update();

            doc.Save(ArtifactsDir + "Field.InsertFieldUsingFieldBuilder.docx");
            //ExEnd:InsertFieldUsingFieldBuilder
        }

        [Test]
        public void InsertAuthorField()
        {
            //ExStart:InsertAuthorField
            //GistId:1cf07762df56f15067d6aef90b14b3db
            Document doc = new Document();

            Paragraph para = (Paragraph) doc.GetChild(NodeType.Paragraph, 0, true);

            // We want to insert an AUTHOR field like this:
            // { AUTHOR Test1 }
            FieldAuthor field = (FieldAuthor) para.AppendField(FieldType.FieldAuthor, false);
            field.AuthorName = "Test1";

            field.Update();

            doc.Save(ArtifactsDir + "WorkingWithFields.InsertAuthorField.docx");
            //ExEnd:InsertAuthorField
        }

        [Test]
        public void InsertAskFieldWithoutDocumentBuilder()
        {
            //ExStart:InsertAskFieldWithoutDocumentBuilder
            //GistId:1cf07762df56f15067d6aef90b14b3db
            Document doc = new Document();

            Paragraph para = (Paragraph) doc.GetChild(NodeType.Paragraph, 0, true);
            // We want to insert an Ask field like this:
            // { ASK \"Test 1\" Test2 \\d Test3 \\o }
            FieldAsk field = (FieldAsk) para.AppendField(FieldType.FieldAsk, false);
            // { ASK \"Test 1\" " }
            field.BookmarkName = "Test 1";
            // { ASK \"Test 1\" Test2 }
            field.PromptText = "Test2";
            // { ASK \"Test 1\" Test2 \\d Test3 }
            field.DefaultResponse = "Test3";
            // { ASK \"Test 1\" Test2 \\d Test3 \\o }
            field.PromptOnceOnMailMerge = true;

            field.Update();

            doc.Save(ArtifactsDir + "WorkingWithFields.InsertAskFieldWithoutDocumentBuilder.docx");
            //ExEnd:InsertAskFieldWithoutDocumentBuilder
        }

        [Test]
        public void InsertAdvanceFieldWithoutDocumentBuilder()
        {
            //ExStart:InsertAdvanceFieldWithoutDocumentBuilder
            //GistId:1cf07762df56f15067d6aef90b14b3db
            Document doc = new Document();

            Paragraph para = (Paragraph) doc.GetChild(NodeType.Paragraph, 0, true);
            // We want to insert an Advance field like this:
            // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 \\y 100 }
            FieldAdvance field = (FieldAdvance) para.AppendField(FieldType.FieldAdvance, false);
            // { ADVANCE \\d 10 " }
            field.DownOffset = "10";
            // { ADVANCE \\d 10 \\l 10 }
            field.LeftOffset = "10";
            // { ADVANCE \\d 10 \\l 10 \\r -3.3 }
            field.RightOffset = "-3.3";
            // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 }
            field.UpOffset = "0";
            // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 }
            field.HorizontalPosition = "100";
            // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 \\y 100 }
            field.VerticalPosition = "100";

            field.Update();

            doc.Save(ArtifactsDir + "WorkingWithFields.InsertAdvanceFieldWithoutDocumentBuilder.docx");
            //ExEnd:InsertAdvanceFieldWithoutDocumentBuilder
        }

        [Test]
        public void GetMailMergeFieldNames()
        {
            //ExStart:GetFieldNames
            //GistId:b4bab1bf22437a86d8062e91cf154494
            Document doc = new Document();

            string[] fieldNames = doc.MailMerge.GetFieldNames();
            //ExEnd:GetFieldNames
            Console.WriteLine("\nDocument have " + fieldNames.Length + " fields.");
        }

        [Test]
        public void MappedDataFields()
        {
            //ExStart:MappedDataFields
            //GistId:b4bab1bf22437a86d8062e91cf154494
            Document doc = new Document();

            doc.MailMerge.MappedDataFields.Add("MyFieldName_InDocument", "MyFieldName_InDataSource");
            //ExEnd:MappedDataFields
        }

        [Test]
        public void DeleteFields()
        {
            //ExStart:DeleteFields
            //GistId:f39874821cb317d245a769c9ce346fea
            Document doc = new Document();

            doc.MailMerge.DeleteFields();
            //ExEnd:DeleteFields
        }

        [Test]
        public void FieldUpdateCulture()
        {
            //ExStart:FieldUpdateCulture
            //GistId:79b46682fbfd7f02f64783b163ed95fc
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertField(FieldType.FieldTime, true);

            doc.FieldOptions.FieldUpdateCultureSource = FieldUpdateCultureSource.FieldCode;
            doc.FieldOptions.FieldUpdateCultureProvider = new FieldUpdateCultureProvider();

            doc.Save(ArtifactsDir + "WorkingWithFields.FieldUpdateCulture.pdf");
            //ExEnd:FieldUpdateCulture
        }

        //ExStart:FieldUpdateCultureProvider
        //GistId:79b46682fbfd7f02f64783b163ed95fc
        class FieldUpdateCultureProvider : IFieldUpdateCultureProvider
        {
            public CultureInfo GetCulture(string name, Field field)
            {
                switch (name)
                {
                    case "ru-RU":
                        CultureInfo culture = new CultureInfo(name, false);
                        DateTimeFormatInfo format = culture.DateTimeFormat;

                        format.MonthNames = new[]
                        {
                            "месяц 1", "месяц 2", "месяц 3", "месяц 4", "месяц 5", "месяц 6", "месяц 7", "месяц 8",
                            "месяц 9", "месяц 10", "месяц 11", "месяц 12", ""
                        };
                        format.MonthGenitiveNames = format.MonthNames;
                        format.AbbreviatedMonthNames = new[]
                        {
                            "мес 1", "мес 2", "мес 3", "мес 4", "мес 5", "мес 6", "мес 7", "мес 8", "мес 9", "мес 10",
                            "мес 11", "мес 12", ""
                        };
                        format.AbbreviatedMonthGenitiveNames = format.AbbreviatedMonthNames;

                        format.DayNames = new[]
                        {
                            "день недели 7", "день недели 1", "день недели 2", "день недели 3", "день недели 4",
                            "день недели 5", "день недели 6"
                        };
                        format.AbbreviatedDayNames = new[]
                            { "день 7", "день 1", "день 2", "день 3", "день 4", "день 5", "день 6" };
                        format.ShortestDayNames = new[] { "д7", "д1", "д2", "д3", "д4", "д5", "д6" };

                        format.AMDesignator = "До полудня";
                        format.PMDesignator = "После полудня";

                        const string pattern = "yyyy MM (MMMM) dd (dddd) hh:mm:ss tt";
                        format.LongDatePattern = pattern;
                        format.LongTimePattern = pattern;
                        format.ShortDatePattern = pattern;
                        format.ShortTimePattern = pattern;

                        return culture;
                    case "en-US":
                        return new CultureInfo(name, false);
                    default:
                        return null;
                }
            }
        }
        //ExEnd:FieldUpdateCultureProvider

        [Test]
        public void FieldDisplayResults()
        {
            //ExStart:FieldDisplayResults
            //GistId:bf0f8a6b40b69a5274ab3553315e147f
            //ExStart:UpdateDocFields
            //GistId:08db64c4d86842c4afd1ecb925ed07c4
            Document document = new Document(MyDir + "Various fields.docx");

            document.UpdateFields();
            //ExEnd:UpdateDocFields

            foreach (Field field in document.Range.Fields)
                Console.WriteLine(field.DisplayResult);
            //ExEnd:FieldDisplayResults
        }

        [Test]
        public void EvaluateIfCondition()
        {
            //ExStart:EvaluateIfCondition
            //GistId:79b46682fbfd7f02f64783b163ed95fc
            DocumentBuilder builder = new DocumentBuilder();

            FieldIf field = (FieldIf) builder.InsertField("IF 1 = 1", null);
            FieldIfComparisonResult actualResult = field.EvaluateCondition();

            Console.WriteLine(actualResult);
            //ExEnd:EvaluateIfCondition
        }

        [Test]
        public void UnlinkFieldsInParagraph()
        {
            //ExStart:UnlinkFieldsInParagraph
            //GistId:f3592014d179ecb43905e37b2a68bc92
            Document doc = new Document(MyDir + "Linked fields.docx");

            // Pass the appropriate parameters to convert all IF fields to text that are encountered only in the last 
            // paragraph of the document.
            doc.FirstSection.Body.LastParagraph.Range.Fields.Where(f => f.Type == FieldType.FieldIf).ToList()
                .ForEach(f => f.Unlink());

            doc.Save(ArtifactsDir + "WorkingWithFields.UnlinkFieldsInParagraph.docx");
            //ExEnd:UnlinkFieldsInParagraph
        }

        [Test]
        public void UnlinkFieldsInDocument()
        {
            //ExStart:UnlinkFieldsInDocument
            //GistId:f3592014d179ecb43905e37b2a68bc92
            Document doc = new Document(MyDir + "Linked fields.docx");

            // Pass the appropriate parameters to convert all IF fields encountered in the document (including headers and footers) to text.
            doc.Range.Fields.Where(f => f.Type == FieldType.FieldIf).ToList().ForEach(f => f.Unlink());

            // Save the document with fields transformed to disk
            doc.Save(ArtifactsDir + "WorkingWithFields.UnlinkFieldsInDocument.docx");
            //ExEnd:UnlinkFieldsInDocument
        }

        [Test]
        public void UnlinkFieldsInBody()
        {
            //ExStart:UnlinkFieldsInBody
            //GistId:f3592014d179ecb43905e37b2a68bc92
            Document doc = new Document(MyDir + "Linked fields.docx");

            // Pass the appropriate parameters to convert PAGE fields encountered to text only in the body of the first section.
            doc.FirstSection.Body.Range.Fields.Where(f => f.Type == FieldType.FieldPage).ToList().ForEach(f => f.Unlink());

            doc.Save(ArtifactsDir + "WorkingWithFields.UnlinkFieldsInBody.docx");
            //ExEnd:UnlinkFieldsInBody
        }

        [Test]
        public void ChangeLocale()
        {
            //ExStart:ChangeLocale
            //GistId:9e90defe4a7bcafb004f73a2ef236986
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertField("MERGEFIELD Date");

            // Store the current culture so it can be set back once mail merge is complete.
            CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;
            // Set to German language so dates and numbers are formatted using this culture during mail merge.
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");

            doc.MailMerge.Execute(new[] { "Date" }, new object[] { DateTime.Now });
            
            Thread.CurrentThread.CurrentCulture = currentCulture;
            
            doc.Save(ArtifactsDir + "WorkingWithFields.ChangeLocale.docx");
            //ExEnd:ChangeLocale
        }

        //ExStart:ConvertFieldsToStaticText
        //GistId:f3592014d179ecb43905e37b2a68bc92
        /// <summary>
        /// Converts any fields of the specified type found in the descendants of the node into static text.
        /// </summary>
        /// <param name="compositeNode">The node in which all descendants of the specified FieldType will be converted to static text.</param>
        /// <param name="targetFieldType">The FieldType of the field to convert to static text.</param>
        private void ConvertFieldsToStaticText(CompositeNode compositeNode, FieldType targetFieldType)
        {
            compositeNode.Range.Fields.Cast<Field>().Where(f => f.Type == targetFieldType).ToList().ForEach(f => f.Unlink());
        }
        //ExEnd:ConvertFieldsToStaticText

        [Test]
        public void FieldResultFormatting()
        {
            //ExStart:FieldResultFormatting
            //GistId:79b46682fbfd7f02f64783b163ed95fc
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            FieldResultFormatter formatter = new FieldResultFormatter("${0}", "Date: {0}", "Item # {0}:");
            doc.FieldOptions.ResultFormatter = formatter;

            // Our field result formatter applies a custom format to newly created fields of three types of formats.
            // Field result formatters apply new formatting to fields as they are updated,
            // which happens as soon as we create them using this InsertField method overload.
            // 1 -  Numeric:
            builder.InsertField(" = 2 + 3 \\# $###");

            Assert.That(doc.Range.Fields[0].Result, Is.EqualTo("$5"));
            Assert.That(formatter.CountFormatInvocations(FieldResultFormatter.FormatInvocationType.Numeric), Is.EqualTo(1));

            // 2 -  Date/time:
            builder.InsertField("DATE \\@ \"d MMMM yyyy\"");

            Assert.That(doc.Range.Fields[1].Result.StartsWith("Date: "), Is.True);
            Assert.That(formatter.CountFormatInvocations(FieldResultFormatter.FormatInvocationType.DateTime), Is.EqualTo(1));

            // 3 -  General:
            builder.InsertField("QUOTE \"2\" \\* Ordinal");

            Assert.That(doc.Range.Fields[2].Result, Is.EqualTo("Item # 2:"));
            Assert.That(formatter.CountFormatInvocations(FieldResultFormatter.FormatInvocationType.General), Is.EqualTo(1));

            formatter.PrintFormatInvocations();
            //ExEnd:FieldResultFormatting
        }

        //ExStart:FieldResultFormatter
        //GistId:79b46682fbfd7f02f64783b163ed95fc
        /// <summary>
        /// When fields with formatting are updated, this formatter will override their formatting
        /// with a custom format, while tracking every invocation.
        /// </summary>
        private class FieldResultFormatter : IFieldResultFormatter
        {
            public FieldResultFormatter(string numberFormat, string dateFormat, string generalFormat)
            {
                mNumberFormat = numberFormat;
                mDateFormat = dateFormat;
                mGeneralFormat = generalFormat;
            }

            public string FormatNumeric(double value, string format)
            {
                if (string.IsNullOrEmpty(mNumberFormat))
                    return null;

                string newValue = String.Format(mNumberFormat, value);
                FormatInvocations.Add(new FormatInvocation(FormatInvocationType.Numeric, value, format, newValue));
                return newValue;
            }

            public string FormatDateTime(DateTime value, string format, CalendarType calendarType)
            {
                if (string.IsNullOrEmpty(mDateFormat))
                    return null;

                string newValue = String.Format(mDateFormat, value);
                FormatInvocations.Add(new FormatInvocation(FormatInvocationType.DateTime, $"{value} ({calendarType})", format, newValue));
                return newValue;
            }

            public string Format(string value, GeneralFormat format)
            {
                return Format((object)value, format);
            }

            public string Format(double value, GeneralFormat format)
            {
                return Format((object)value, format);
            }

            private string Format(object value, GeneralFormat format)
            {
                if (string.IsNullOrEmpty(mGeneralFormat))
                    return null;

                string newValue = String.Format(mGeneralFormat, value);
                FormatInvocations.Add(new FormatInvocation(FormatInvocationType.General, value, format.ToString(), newValue));
                return newValue;
            }

            public int CountFormatInvocations(FormatInvocationType formatInvocationType)
            {
                if (formatInvocationType == FormatInvocationType.All)
                    return FormatInvocations.Count;
                return FormatInvocations.Count(f => f.FormatInvocationType == formatInvocationType);
            }

            public void PrintFormatInvocations()
            {
                foreach (FormatInvocation f in FormatInvocations)
                    Console.WriteLine($"Invocation type:\t{f.FormatInvocationType}\n" +
                                      $"\tOriginal value:\t\t{f.Value}\n" +
                                      $"\tOriginal format:\t{f.OriginalFormat}\n" +
                                      $"\tNew value:\t\t\t{f.NewValue}\n");
            }

            private readonly string mNumberFormat;
            private readonly string mDateFormat;
            private readonly string mGeneralFormat;
            private List<FormatInvocation> FormatInvocations { get; } = new List<FormatInvocation>();

            private class FormatInvocation
            {
                public FormatInvocationType FormatInvocationType { get; }
                public object Value { get; }
                public string OriginalFormat { get; }
                public string NewValue { get; }

                public FormatInvocation(FormatInvocationType formatInvocationType, object value, string originalFormat, string newValue)
                {
                    Value = value;
                    FormatInvocationType = formatInvocationType;
                    OriginalFormat = originalFormat;
                    NewValue = newValue;
                }
            }

            public enum FormatInvocationType
            {
                Numeric, DateTime, General, All
            }
        }
        //ExEnd:FieldResultFormatter
    }
}