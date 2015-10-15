using Aspose.Words;
using Aspose.Words.Fields;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Obtaining_Form_Fields
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = @"E:\Aspose\Aspose Vs OpenXML\Missing Features of OpenXML Words Provided by Aspose.Words v1.1\Sample Files\";

            //Shows how to get a collection of form fields.
            Document doc = new Document(MyDir + "FormFields.doc");
            FormFieldCollection formFields = doc.Range.FormFields;


            //Shows how to access form fields.
            Document myDoc = new Document(MyDir + "FormFields.doc");
            FormFieldCollection documentFormFields = myDoc.Range.FormFields;

            FormField formField1 = documentFormFields[3];
            FormField formField2 = documentFormFields["CustomerName"];

        }
    }
}
