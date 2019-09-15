using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Loading_and_Saving
{
    class WorkingWithVbaMacros
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            ReadVbaMacros(dataDir);
        }

        public static void ReadVbaMacros(string dataDir)
        {
            //ExStart:ReadVbaMacros
            Document doc = new Document(dataDir + "Document.dot");

            if (doc.VbaProject != null)
            {
                foreach (VbaModule module in doc.VbaProject.Modules)
                {
                    Console.WriteLine(module.SourceCode);
                }
            }
            //ExEnd:ReadVbaMacros
        }

        public static void ModifyVbaMacros(string dataDir)
        {
            //ExStart:ModifyVbaMacros
            Document doc = new Document(dataDir + "test.docm");
            VbaProject project = doc.VbaProject;

            const string newSourceCode = "Test change source code";

            // Choose a module, and set a new source code.
            project.Modules[0].SourceCode = newSourceCode;
            //ExEnd:ModifyVbaMacros
        }
    }
}
