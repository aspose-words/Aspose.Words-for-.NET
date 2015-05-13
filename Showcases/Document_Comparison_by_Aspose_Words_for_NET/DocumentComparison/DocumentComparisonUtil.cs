using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;

namespace DocumentComparison
{
    /// <summary>
    /// Class to compare two MS Word documents
    /// </summary>
    public class DocumentComparisonUtil
    {
        

        /// <summary>
        /// Compare the two documents using Aspose.Words and save the result as a Word document
        /// </summary>
        /// <param name="document1">First document</param>
        /// <param name="document2">Second document</param>
        /// <param name="comparisonDocument">Comparison document</param>
        public void Compare(string document1, string document2, string comparisonDocument, ref int added, ref int deleted)
        {
            added = 0;
            deleted = 0;

            // Load both documents in Aspose.Words
            Document doc1 = new Document(document1);
            Document doc2 = new Document(document2);
            Document docComp = new Document(document1);
            DocumentBuilder builder = new DocumentBuilder(docComp);

            doc1.Compare(doc2, "a", DateTime.Now);

            foreach (Revision revision in doc1.Revisions)
            {
                switch (revision.RevisionType)
                {
                    case RevisionType.Insertion:
                        added++;
                        break;
                    case RevisionType.Deletion:
                        deleted++;
                        break;
                }
                Console.WriteLine(revision.RevisionType + ": " + revision.ParentNode);
            }

            Debug.WriteLine("Revisions: " + doc1.Revisions.Count);
            doc1.Save(comparisonDocument);
        }
    }

}