using Aspose.Words.Drawing;
using Aspose.Words.MailMerging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Mail_Merge
{
    class MailMergeImageField
    {
        public static void Run()
        {
            // ExStart:MailMergeImageField       
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting();
            Document doc = new Document(dataDir + "template.docx");

            doc.MailMerge.UseNonMergeFields = true;
            doc.MailMerge.TrimWhitespaces = true;
            doc.MailMerge.UseWholeParagraphAsRegion = false;
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyTableRows
                    | MailMergeCleanupOptions.RemoveContainingFields
                    | MailMergeCleanupOptions.RemoveUnusedRegions
                    | MailMergeCleanupOptions.RemoveUnusedFields;

            // Add a handler for the MergeField event.
            doc.MailMerge.FieldMergingCallback = new ImageFieldMergingHandler();
            doc.MailMerge.ExecuteWithRegions(new DataSourceRoot());

            dataDir = dataDir + "MailMerge.ImageMailMerge_out.doc";
            doc.Save(dataDir);
            // ExEnd:MailMergeImageField
            Console.WriteLine("\nMail merge Image Field performed successfully.\nFile saved at " + dataDir);
        }
        // ExStart:ImageFieldMergingHandler
        private class ImageFieldMergingHandler : IFieldMergingCallback
        {
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
            {
                //  Implementation is not required.
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                Shape shape = new Shape(args.Document, ShapeType.Image);
                shape.Width = 126;
                shape.Height = 126;
                shape.WrapType = WrapType.Square;

                string imageFileName = Path.GetFullPath(RunExamples.GetDataDir_WorkingWithDocument() + "image.png");
                shape.ImageData.SetImage(imageFileName);

                args.Shape = shape;
            }
        }
        // ExEnd:ImageFieldMergingHandler
        // ExStart:DataSourceRoot
        public class DataSourceRoot : IMailMergeDataSourceRoot
        {
            public IMailMergeDataSource GetDataSource(String s)
            {
                return new DataSource();
            }

            private class DataSource : IMailMergeDataSource
            {

                bool next = true;

                string IMailMergeDataSource.TableName => TableName();

                public string TableName()
                {
                    return "example";
                }

                public bool MoveNext()
                {
                    bool result = next;
                    next = false;
                    return result;
                }

                public IMailMergeDataSource GetChildDataSource(String s)
                {
                    return null;
                }

                public bool GetValue(string fieldName, out object fieldValue)
                {
                    fieldValue = null;
                    return false;
                }
            }
        }
        // ExEnd:DataSourceRoot
    }
}
