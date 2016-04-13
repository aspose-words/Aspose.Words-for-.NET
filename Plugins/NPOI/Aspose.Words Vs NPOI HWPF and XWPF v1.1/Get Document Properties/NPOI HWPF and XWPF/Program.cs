using NPOI.HPSF;
using System;
using System.IO;

namespace NPOI_HWPF_and_XWPF
{
    class Program
    {
        static void Main(string[] args)
        {
            // NPOI library doest not have ablitity to read word document. 

            SummaryInformation summaryInfo = new SummaryInformation(new PropertySet(new FileStream("data/Get Document Properties.doc", FileMode.Open)));
            Console.WriteLine(summaryInfo.ApplicationName);
            Console.WriteLine(summaryInfo.Author);
            Console.WriteLine(summaryInfo.Comments);
            Console.WriteLine(summaryInfo.CharCount);
            Console.WriteLine(summaryInfo.EditTime);
            Console.WriteLine(summaryInfo.Keywords);
            Console.WriteLine(summaryInfo.LastAuthor);
            Console.WriteLine(summaryInfo.PageCount);
            Console.WriteLine(summaryInfo.RevNumber);
            Console.WriteLine(summaryInfo.Security);
            Console.WriteLine(summaryInfo.Subject);
            Console.WriteLine(summaryInfo.Template);
        }
    }
}
