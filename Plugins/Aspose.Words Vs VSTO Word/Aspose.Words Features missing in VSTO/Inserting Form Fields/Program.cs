using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Inserting_Form_Fields
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            string[] items = { "One", "Two", "Three" };
            builder.InsertComboBox("DropDown", items, 0);

            
        }
    }
}
