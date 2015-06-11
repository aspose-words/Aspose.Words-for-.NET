using System.Reflection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace CSharp.Quick_Start
{
    class _RunExamples
    {
        public static void Main()
        {
            // Run the examples. Un-comment the one you want to run
            AppendDocuments.Run();
            ApplyLicense.Run();
            Doc2Pdf.Run();
            FindAndReplace.Run();
            HelloWorld.Run();
            LoadAndSaveToDisk.Run();
            LoadAndSaveToStream.Run();
            SimpleMailMerge.Run();
            UpdateFields.Run();
            WorkingWithNodes.Run();

            // Stop before exiting
            Console.WriteLine("\n\nProgram Finished. Press any key to exit....");
            Console.ReadKey();
        }

        public static String GetDataDir()
        {
            return Path.GetFullPath("../../Quick-Start/Data/");
        }
    }
}
