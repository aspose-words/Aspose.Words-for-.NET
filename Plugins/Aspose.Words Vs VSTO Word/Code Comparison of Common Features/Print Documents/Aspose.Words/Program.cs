using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
            string FileName = "YourFileName.docx";
            Document doc = new Document(FileName);
            doc.Print();
        }
    }
}
