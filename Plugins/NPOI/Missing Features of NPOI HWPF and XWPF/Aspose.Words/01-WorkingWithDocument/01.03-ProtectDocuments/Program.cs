using System;
using System.Collections.Generic;
using System.Text; using Aspose.Words;

namespace _01._03_ProtectDocuments
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document("../../data/document.doc");
            doc.Protect(ProtectionType.ReadOnly);

            // Following other Protection types are also available
            // ProtectionType.NoProtection
            // ProtectionType.AllowOnlyRevisions
            // ProtectionType.AllowOnlyComments
            // ProtectionType.AllowOnlyFormFields

            doc.Save("AsposeProtect.doc", SaveFormat.Doc);
        }
    }
}
