using Aspose.Words;
using Aspose.Words.Fields;

namespace DocsExamples.Mail_Merge_and_Reporting.Custom_examples
{
    internal class CreateMailMergeTemplate
    {
        //ExStart:CreateMailMergeTemplate
        public Document Template()
        {
            DocumentBuilder builder = new DocumentBuilder();

            // Insert a text input field the unique name of this field is "Hello", the other parameters define
            // what type of FormField it is, the format of the text, the field result and the maximum text length (0 = no limit)
            builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", "Hello", 0);
            builder.InsertField(@"MERGEFIELD CustomerFirstName \* MERGEFORMAT");

            builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "", " ", 0);
            builder.InsertField(@"MERGEFIELD CustomerLastName \* MERGEFORMAT");

            builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "", " , ", 0);

            // Inserts a paragraph break into the document
            builder.InsertParagraph();

            // Insert mail body
            builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", "Thanks for purchasing our ", 0);
            builder.InsertField(@"MERGEFIELD ProductName \* MERGEFORMAT");

            builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", ", please download your Invoice at ",
                0);
            builder.InsertField(@"MERGEFIELD InvoiceURL \* MERGEFORMAT");

            builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "",
                ". If you have any questions please call ", 0);
            builder.InsertField(@"MERGEFIELD Supportphone \* MERGEFORMAT");

            builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", ", or email us at ", 0);
            builder.InsertField(@"MERGEFIELD SupportEmail \* MERGEFORMAT");

            builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", ".", 0);

            builder.InsertParagraph();

            // Insert mail ending
            builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", "Best regards,", 0);
            builder.InsertBreak(BreakType.LineBreak);
            builder.InsertField(@"MERGEFIELD EmployeeFullname \* MERGEFORMAT");

            builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "", " ", 0);
            builder.InsertField(@"MERGEFIELD EmployeeDepartment \* MERGEFORMAT");

            return builder.Document;
        }
        //ExEnd:CreateMailMergeTemplate
    }
}
