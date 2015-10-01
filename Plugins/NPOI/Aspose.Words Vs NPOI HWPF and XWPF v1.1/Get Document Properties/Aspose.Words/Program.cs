using Aspose.Words.Properties;
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
            Document doc = new Document("data/Get Document Properties.doc");
            foreach (DocumentProperty prop in doc.BuiltInDocumentProperties)
            {
                Console.WriteLine(prop.Name+": "+ prop.Value);

            }
        }
    }
}
