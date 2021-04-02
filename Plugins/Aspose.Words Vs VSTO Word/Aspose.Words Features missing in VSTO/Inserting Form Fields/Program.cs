using Aspose.Words;

namespace Inserting_Form_Fields
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a drop down combo box with three options for the user to select.
            string[] items = { "One", "Two", "Three" };
            builder.InsertComboBox("DropDown", items, 0);
        }
    }
}
