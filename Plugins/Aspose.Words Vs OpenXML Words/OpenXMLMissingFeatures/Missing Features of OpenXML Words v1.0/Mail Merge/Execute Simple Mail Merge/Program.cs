// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Words;

namespace Execute_Simple_Mail_Merge
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = @"Files\";
            // Open an existing document.
            Document doc = new Document(MyDir + "Merge Field.doc");

            // Fill the fields in the document with user data.
            doc.MailMerge.Execute(
                new string[] { "Name", "City" },
                new object[] { "Zeeshan", "Islamabad" });

            // Send the document in Word format to the client browser with an option to save to disk or open inside the current browser.
            doc.Save(MyDir + "MailMerge.ExecuteArray Out.doc");
        }
    }
}
