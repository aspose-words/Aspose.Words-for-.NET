using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents
{
    internal class WorkingWithFields : DocsExamplesBase
    {
        [Test]
        public void ChangeFieldUpdateCultureSource()
        {
            //ExStart:ChangeFieldUpdateCultureSource
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
            //ExStart:SpecifylocaleAtFieldlevel
            DocumentBuilder builder = new DocumentBuilder();

            Field field = builder.InsertField(FieldType.FieldDate, true);
            field.LocaleId = 1049;
            
            builder.Document.Save(ArtifactsDir + "WorkingWithFields.SpecifylocaleAtFieldlevel.docx");
            //ExEnd:SpecifylocaleAtFieldlevel
        }

        [Test]
        public void ReplaceHyperlinks()
        {
            //ExStart:ReplaceHyperlinks
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
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertField(@"MERGEFIELD MyMergeField1 \* MERGEFORMAT");
            builder.InsertField(@"MERGEFIELD MyMergeField2 \* MERGEFORMAT");

            // Select all field start nodes so we can find the merge fields.
            NodeCollection fieldStarts = doc.GetChildNodes(NodeType.FieldStart, true);
            foreach (FieldStart fieldStart in fieldStarts)
            {
                if (fieldStart.FieldType == FieldType.FieldMergeField)
                {
                    MergeField mergeField = new MergeField(fieldStart);
                    mergeField.Name += "_Renamed";
                }
            }

            doc.Save(ArtifactsDir + "WorkingWithFields.RenameMergeFields.doc");
            //ExEnd:RenameMergeFields
        }

        //ExStart:MergeField
        /// <summary>
        /// Represents a facade object for a merge field in a Microsoft Word document.
        /// </summary>
        internal class MergeField
        {
            internal MergeField(FieldStart fieldStart)
            {
                if (fieldStart == null)
                    throw new ArgumentNullException(nameof(fieldStart));
                if (fieldStart.FieldType != FieldType.FieldMergeField)
                    throw new ArgumentException("Field start type must be FieldMergeField.");

                mFieldStart = fieldStart;

                // Find the field separator node.
                mFieldSeparator = fieldStart.GetField().Separator;
                if (mFieldSeparator == null)
                    throw new InvalidOperationException("Cannot find field separator.");

                mFieldEnd = fieldStart.GetField().End;
            }

            /// <summary>
            /// Gets or sets the name of the merge field.
            /// </summary>
            internal string Name
            {
                get => ((FieldStart) mFieldStart).GetField().Result.Replace("«", "").Replace("»", "");
                set
                {
                    // Merge field name is stored in the field result which is a Run
                    // node between field separator and field end.
                    Run fieldResult = (Run) mFieldSeparator.NextSibling;
                    fieldResult.Text = string.Format("«{0}»", value);

                    // But sometimes the field result can consist of more than one run, delete these runs.
                    RemoveSameParent(fieldResult.NextSibling, mFieldEnd);

                    UpdateFieldCode(value);
                }
            }

            private void UpdateFieldCode(string fieldName)
            {
                // Field code is stored in a Run node between field start and field separator.
                Run fieldCode = (Run) mFieldStart.NextSibling;

                Match match = gRegex.Match(((FieldStart) mFieldStart).GetField().GetFieldCode());

                string newFieldCode = string.Format(" {0}{1} ", match.Groups["start"].Value, fieldName);
                fieldCode.Text = newFieldCode;

                // But sometimes the field code can consist of more than one run, delete these runs.
                RemoveSameParent(fieldCode.NextSibling, mFieldSeparator);
            }

            /// <summary>
            /// Removes nodes from start up to but not including the end node.
            /// Start and end are assumed to have the same parent.
            /// </summary>
            private void RemoveSameParent(Node startNode, Node endNode)
            {
                if (endNode != null && startNode.ParentNode != endNode.ParentNode)
                    throw new ArgumentException("Start and end nodes are expected to have the same parent.");

                Node curChild = startNode;
                while (curChild != null && curChild != endNode)
                {
                    Node nextChild = curChild.NextSibling;
                    curChild.Remove();
                    curChild = nextChild;
                }
            }

            private readonly Node mFieldStart;
            private readonly Node mFieldSeparator;
            private readonly Node mFieldEnd;

            private readonly Regex gRegex = new Regex(@"\s*(?<start>MERGEFIELD\s|)(\s|)(?<name>\S+)\s+");
        }
        //ExEnd:MergeField

        [Test]
        public void RemoveField()
        {
            //ExStart:RemoveField
            Document doc = new Document(MyDir + "Various fields.docx");
            
            Field field = doc.Range.Fields[0];
            field.Remove();
            //ExEnd:RemoveField
        }

        [Test]
        public void InsertTOAFieldWithoutDocumentBuilder()
        {
            //ExStart:InsertTOAFieldWithoutDocumentBuilder
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

            doc.Save(ArtifactsDir + "WorkingWithFields.InsertTOAFieldWithoutDocumentBuilder.docx");
            //ExEnd:InsertTOAFieldWithoutDocumentBuilder
        }

        [Test]
        public void InsertNestedFields()
        {
            //ExStart:InsertNestedFields
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
        public void InsertMergeFieldUsingDOM()
        {
            //ExStart:InsertMergeFieldUsingDOM
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Paragraph para = (Paragraph) doc.GetChildNodes(NodeType.Paragraph, true)[0];

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

            // Finally update this merge field
            field.Update();

            doc.Save(ArtifactsDir + "WorkingWithFields.InsertMergeFieldUsingDOM.docx");
            //ExEnd:InsertMergeFieldUsingDOM
        }

        [Test]
        public void InsertMailMergeAddressBlockFieldUsingDOM()
        {
            //ExStart:InsertMailMergeAddressBlockFieldUsingDOM
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Paragraph para = (Paragraph) doc.GetChildNodes(NodeType.Paragraph, true)[0];

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

            doc.Save(ArtifactsDir + "WorkingWithFields.InsertMailMergeAddressBlockFieldUsingDOM.docx");
            //ExEnd:InsertMailMergeAddressBlockFieldUsingDOM
        }

        [Test]
        public void InsertFieldIncludeTextWithoutDocumentBuilder()
        {
            //ExStart:InsertFieldIncludeTextWithoutDocumentBuilder
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
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.InsertField(@"MERGEFIELD MyFieldName \* MERGEFORMAT");
            
            doc.Save(ArtifactsDir + "WorkingWithFields.InsertField.docx");
            //ExEnd:InsertField
        }

        [Test]
        public void InsertAuthorField()
        {
            //ExStart:InsertAuthorField
            Document doc = new Document();

            Paragraph para = (Paragraph) doc.GetChildNodes(NodeType.Paragraph, true)[0];

            // We want to insert an AUTHOR field like this:
            // { AUTHOR Test1 }

            FieldAuthor field = (FieldAuthor) para.AppendField(FieldType.FieldAuthor, false);            
            field.AuthorName = "Test1"; // { AUTHOR Test1 }

            field.Update();

            doc.Save(ArtifactsDir + "WorkingWithFields.InsertAuthorField.docx");
            //ExEnd:InsertAuthorField
        }

        [Test]
        public void InsertASKFieldWithOutDocumentBuilder()
        {
            //ExStart:InsertASKFieldWithOutDocumentBuilder
            Document doc = new Document();

            Paragraph para = (Paragraph) doc.GetChildNodes(NodeType.Paragraph, true)[0];

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

            doc.Save(ArtifactsDir + "WorkingWithFields.InsertASKFieldWithOutDocumentBuilder.docx");
            //ExEnd:InsertASKFieldWithOutDocumentBuilder
        }

        [Test]
        public void InsertAdvanceFieldWithOutDocumentBuilder()
        {
            //ExStart:InsertAdvanceFieldWithOutDocumentBuilder
            Document doc = new Document();

            Paragraph para = (Paragraph) doc.GetChildNodes(NodeType.Paragraph, true)[0];

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

            doc.Save(ArtifactsDir + "WorkingWithFields.InsertAdvanceFieldWithOutDocumentBuilder.docx");
            //ExEnd:InsertAdvanceFieldWithOutDocumentBuilder
        }

        [Test]
        public void GetMailMergeFieldNames()
        {
            //ExStart:GetFieldNames
            Document doc = new Document();

            string[] fieldNames = doc.MailMerge.GetFieldNames();
            //ExEnd:GetFieldNames
            Console.WriteLine("\nDocument have " + fieldNames.Length + " fields.");
        }

        [Test]
        public void MappedDataFields()
        {
            //ExStart:MappedDataFields
            Document doc = new Document();

            doc.MailMerge.MappedDataFields.Add("MyFieldName_InDocument", "MyFieldName_InDataSource");
            //ExEnd:MappedDataFields
        }

        [Test]
        public void DeleteFields()
        {
            //ExStart:DeleteFields
            Document doc = new Document();

            doc.MailMerge.DeleteFields();
            //ExEnd:DeleteFields
        }

        [Test]
        public void FieldUpdateCulture()
        {
            //ExStart:FieldUpdateCultureProvider
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertField(FieldType.FieldTime, true);

            doc.FieldOptions.FieldUpdateCultureSource = FieldUpdateCultureSource.FieldCode;
            doc.FieldOptions.FieldUpdateCultureProvider = new FieldUpdateCultureProvider();

            doc.Save(ArtifactsDir + "WorkingWithFields.FieldUpdateCulture.pdf");
            //ExEnd:FieldUpdateCultureProvider
        }

        //ExStart:FieldUpdateCultureProviderGetCulture
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
        //ExEnd:FieldUpdateCultureProviderGetCulture

        [Test]
        public void FieldDisplayResults()
        {
            //ExStart:FieldDisplayResults
            //ExStart:UpdateDocFields
            Document document = new Document(MyDir + "Various fields.docx");

            document.UpdateFields();
            //ExEnd:UpdateDocFields

            foreach (Field field in document.Range.Fields)
                Console.WriteLine(field.DisplayResult);
            //ExEnd:FieldDisplayResults
        }

        [Test]
        public void EvaluateIFCondition()
        {
            //ExStart:EvaluateIFCondition
            DocumentBuilder builder = new DocumentBuilder();

            FieldIf field = (FieldIf) builder.InsertField("IF 1 = 1", null);
            FieldIfComparisonResult actualResult = field.EvaluateCondition();

            Console.WriteLine(actualResult);
            //ExEnd:EvaluateIFCondition
        }

        [Test]
        public void ConvertFieldsInParagraph()
        {
            //ExStart:ConvertFieldsInParagraph
            Document doc = new Document(MyDir + "Linked fields.docx");

            // Pass the appropriate parameters to convert all IF fields to text that are encountered only in the last 
            // paragraph of the document.
            doc.FirstSection.Body.LastParagraph.Range.Fields.Where(f => f.Type == FieldType.FieldIf).ToList()
                .ForEach(f => f.Unlink());

            doc.Save(ArtifactsDir + "WorkingWithFields.TestFile.docx");
            //ExEnd:ConvertFieldsInParagraph
        }

        [Test]
        public void ConvertFieldsInDocument()
        {
            //ExStart:ConvertFieldsInDocument
            Document doc = new Document(MyDir + "Linked fields.docx");

            // Pass the appropriate parameters to convert all IF fields encountered in the document (including headers and footers) to text.
            doc.Range.Fields.Where(f => f.Type == FieldType.FieldIf).ToList().ForEach(f => f.Unlink());

            // Save the document with fields transformed to disk
            doc.Save(ArtifactsDir + "WorkingWithFields.ConvertFieldsInDocument.docx");
            //ExEnd:ConvertFieldsInDocument
        }

        [Test]
        public void ConvertFieldsInBody()
        {
            //ExStart:ConvertFieldsInBody
            Document doc = new Document(MyDir + "Linked fields.docx");

            // Pass the appropriate parameters to convert PAGE fields encountered to text only in the body of the first section.
            doc.FirstSection.Body.Range.Fields.Where(f => f.Type == FieldType.FieldPage).ToList().ForEach(f => f.Unlink());

            doc.Save(ArtifactsDir + "WorkingWithFields.ConvertFieldsInBody.docx");
            //ExEnd:ConvertFieldsInBody
        }

        [Test]
        public void ChangeLocale()
        {
            //ExStart:ChangeLocale
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
    }
}