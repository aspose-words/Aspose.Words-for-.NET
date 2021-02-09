using System.Data;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.MailMerging;
using NUnit.Framework;

namespace DocsExamples.Mail_Merge_and_Reporting
{
    internal class WorkingWithCleanupOptions : DocsExamplesBase
    {
        [Test]
        public void RemoveRowsFromTable()
        {
            //ExStart:RemoveRowsFromTable
            Document doc = new Document(MyDir + "Mail merge destination - Northwind suppliers.docx");
            
            DataSet data = new DataSet();
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedRegions |
                                           MailMergeCleanupOptions.RemoveEmptyTableRows;

            doc.MailMerge.MergeDuplicateRegions = true;
            doc.MailMerge.ExecuteWithRegions(data);

            doc.Save(ArtifactsDir + "WorkingWithCleanupOptions.RemoveRowsFromTable.docx");
            //ExEnd:RemoveRowsFromTable
        }

        [Test]
        public void CleanupParagraphsWithPunctuationMarks()
        {
            //ExStart:CleanupParagraphsWithPunctuationMarks
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            FieldMergeField mergeFieldOption1 = (FieldMergeField) builder.InsertField("MERGEFIELD", "Option_1");
            mergeFieldOption1.FieldName = "Option_1";

            // Here is the complete list of cleanable punctuation marks: ! , . : ; ? ¡ ¿.
            builder.Write(" ?  ");

            FieldMergeField mergeFieldOption2 = (FieldMergeField) builder.InsertField("MERGEFIELD", "Option_2");
            mergeFieldOption2.FieldName = "Option_2";

            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
            // The option's default value is true, which means that the behavior was changed to mimic MS Word.
            // If you rely on the old behavior can revert it by setting the option to false.
            doc.MailMerge.CleanupParagraphsWithPunctuationMarks = true;

            doc.MailMerge.Execute(new[] { "Option_1", "Option_2" }, new object[] { null, null });

            doc.Save(ArtifactsDir + "WorkingWithCleanupOptions.CleanupParagraphsWithPunctuationMarks.docx");
            //ExEnd:CleanupParagraphsWithPunctuationMarks
        }

        [Test]
        public void RemoveUnmergedRegions()
        {
            //ExStart:RemoveUnmergedRegions
            Document doc = new Document(MyDir + "Mail merge destination - Northwind suppliers.docx");

            DataSet data = new DataSet();
            //ExStart:MailMergeCleanupOptions
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedRegions;
            // doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveContainingFields;
            // doc.MailMerge.CleanupOptions |= MailMergeCleanupOptions.RemoveStaticFields;
            // doc.MailMerge.CleanupOptions |= MailMergeCleanupOptions.RemoveEmptyParagraphs;           
            // doc.MailMerge.CleanupOptions |= MailMergeCleanupOptions.RemoveUnusedFields;
            //ExEnd:MailMergeCleanupOptions

            // Merge the data with the document by executing mail merge which will have no effect as there is no data.
            // However the regions found in the document will be removed automatically as they are unused.
            doc.MailMerge.ExecuteWithRegions(data);

            doc.Save(ArtifactsDir + "WorkingWithCleanupOptions.RemoveUnmergedRegions.docx");
            //ExEnd:RemoveUnmergedRegions
        }

        [Test]
        public void RemoveEmptyParagraphs()
        {
            //ExStart:RemoveEmptyParagraphs
            Document doc = new Document(MyDir + "Table with fields.docx");

            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;

            doc.MailMerge.Execute(new string[] { "FullName", "Company", "Address", "Address2", "City" },
                new object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });

            doc.Save(ArtifactsDir + "WorkingWithCleanupOptions.RemoveEmptyParagraphs.docx");
            //ExEnd:RemoveEmptyParagraphs
        }

        [Test]
        public void RemoveUnusedFields()
        {
            //ExStart:RemoveUnusedFields
            Document doc = new Document(MyDir + "Table with fields.docx");

            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;

            doc.MailMerge.Execute(new string[] { "FullName", "Company", "Address", "Address2", "City" },
                new object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });

            doc.Save(ArtifactsDir + "WorkingWithCleanupOptions.RemoveUnusedFields.docx");
            //ExEnd:RemoveUnusedFields
        }

        [Test]
        public void RemoveContainingFields()
        {
            //ExStart:RemoveContainingFields
            Document doc = new Document(MyDir + "Table with fields.docx");

            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveContainingFields;

            doc.MailMerge.Execute(new string[] { "FullName", "Company", "Address", "Address2", "City" },
                new object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });

            doc.Save(ArtifactsDir + "WorkingWithCleanupOptions.RemoveContainingFields.docx");
            //ExEnd:RemoveContainingFields
        }

        [Test]
        public void RemoveEmptyTableRows()
        {
            //ExStart:RemoveEmptyTableRows
            Document doc = new Document(MyDir + "Table with fields.docx");

            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyTableRows;

            doc.MailMerge.Execute(new string[] { "FullName", "Company", "Address", "Address2", "City" },
                new object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });

            doc.Save(ArtifactsDir + "WorkingWithCleanupOptions.RemoveEmptyTableRows.docx");
            //ExEnd:RemoveEmptyTableRows
        }
    }
}