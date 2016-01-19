using System;
using System.Collections.Generic;
using System.Text;

namespace Conversion_from_docx_to_doc
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = @"Files\";
            Document doc = new Document(MyDir + "Sample.docx");
            doc.Save(MyDir + "Converted.doc", SaveFormat.Doc);
        }
    }
}
