
using System.IO;
using Aspose.Words;
using System;
using System.Collections;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class GetVariables
    {
        public static void Run()
        {
            //ExStart:GetVariables
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            // Load the template document.
            Document doc = new Document(dataDir + "TestFile.doc");
            string variables = "";
            foreach (DictionaryEntry entry in doc.Variables)
            {
                string name = entry.Key.ToString();
                string value = entry.Value.ToString();
                if (variables == "")
                {
                    // Do something useful.
                    variables = "Name: " + name + "," + "Value: {1}" + value;
                }
                else
                {
                    variables = variables + "Name: " + name + "," + "Value: {1}" + value;
                }
            }
            //ExEnd:GetVariables
            Console.WriteLine("\nDocument have following variables " + variables);
        }
        
    }
}
