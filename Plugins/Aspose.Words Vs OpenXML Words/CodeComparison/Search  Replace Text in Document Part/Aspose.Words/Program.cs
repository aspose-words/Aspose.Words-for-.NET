using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
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
