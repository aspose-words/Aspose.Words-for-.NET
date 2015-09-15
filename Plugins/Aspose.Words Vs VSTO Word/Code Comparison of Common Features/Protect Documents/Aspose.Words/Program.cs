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
            Document doc = new Document();
            doc.Protect(ProtectionType.AllowOnlyFormFields, "password");
        }
    }
}
