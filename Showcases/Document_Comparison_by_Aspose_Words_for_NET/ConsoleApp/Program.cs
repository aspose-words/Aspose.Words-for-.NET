using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp
{
    class Program
    {
        static String dataDir = "../../UserFiles/";

        static void Main(string[] args)
        {
            try
            {
                DocumentComparison.Common.SetLicense();

                string document1 = dataDir + "Paris Trip1.docx";
                string document2 = dataDir + "Paris Trip2.docx";
                string comparisonDocument = GetCompareDocumentName(document1, document2);
                int added = 0, deleted = 0;

                DocumentComparison.DocumentComparisonUtil comp = new DocumentComparison.DocumentComparisonUtil();
                comp.Compare(document1, document2, comparisonDocument, ref added, ref deleted);

                Console.WriteLine("comparison document: " + comparisonDocument);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            Console.WriteLine("Program finished successfully.");
            Console.ReadKey();
        }

        private static string GetCompareDocumentName(string document1, string document2)
        {
            return dataDir + Path.GetFileNameWithoutExtension(document1) + " Compared to " +
                Path.GetFileNameWithoutExtension(document2) + ".docx";
        }
    }
}
