using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document("Test.docx");
            Regex regex = new Regex("Hello World!", RegexOptions.IgnoreCase);
            doc.Range.Replace(regex, "Hi Everyone!");
            doc.Save("Test.docx");
        }
    }
}
