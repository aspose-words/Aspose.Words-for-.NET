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

            CreateVbaProject(dataDir);
            ReadVbaMacros(dataDir);
            ModifyVbaMacros(dataDir);
            CloneVbaProject(dataDir);
            CloneVbaModule(dataDir);
        }

        public static void CreateVbaProject(string dataDir)
        {
            //ExStart:CreateVbaProject
            Document doc = new Document();

            // Create a new VBA project.
            VbaProject project = new VbaProject();
            project.Name = "AsposeProject";
            doc.VbaProject = project;

            // Create a new module and specify a macro source code.
            VbaModule module = new VbaModule();
            module.Name = "AsposeModule";
            module.Type = VbaModuleType.ProceduralModule;
            module.SourceCode = "New source code";

            // Add module to the VBA project.
            doc.VbaProject.Modules.Add(module);

            doc.Save(dataDir + "VbaProject_out.docm");
            //ExEnd:CreateVbaProject
            Console.WriteLine("\nDocument saved successfully.\nFile saved at " + dataDir);
        }

        public static void ReadVbaMacros(string dataDir)
        {
            //ExStart:ReadVbaMacros
            Document doc = new Document(dataDir + "VbaProject_out.docm");

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
            Document doc = new Document(dataDir + "VbaProject_out.docm");
            VbaProject project = doc.VbaProject;

            const string newSourceCode = "Test change source code";

            // Choose a module, and set a new source code.
            project.Modules[0].SourceCode = newSourceCode;

            doc.Save(dataDir + "VbaProject_out.docm");
            //ExEnd:ModifyVbaMacros
            Console.WriteLine("\nDocument saved successfully.\nFile saved at " + dataDir);
        }

        public static void CloneVbaProject(string dataDir)
        {
            //ExStart:CloneVbaProject
            Document doc = new Document(dataDir + "VbaProject_source.docm");
            VbaProject project = doc.VbaProject;

            Document destDoc = new Document();

            // Clone the whole project.
            destDoc.VbaProject = doc.VbaProject.Clone();

            destDoc.Save(dataDir + "output.docm");
            //ExEnd:CloneVbaProject
            Console.WriteLine("\nDocument saved successfully.\nFile saved at " + dataDir);
        }

        public static void CloneVbaModule(string dataDir)
        {
            //ExStart:CloneVbaModule
            Document doc = new Document(dataDir + "VbaProject_source.docm");
            VbaProject project = doc.VbaProject;

            Document destDoc = new Document();

            destDoc.VbaProject = new VbaProject();

            // Clone a single module.
            VbaModule copyModule = doc.VbaProject.Modules["Module1"].Clone();
            destDoc.VbaProject.Modules.Add(copyModule);

            destDoc.Save(dataDir + "output.docm");
            //ExEnd:CloneVbaModule
            Console.WriteLine("\nDocument saved successfully.\nFile saved at " + dataDir);
        }
    }
}
