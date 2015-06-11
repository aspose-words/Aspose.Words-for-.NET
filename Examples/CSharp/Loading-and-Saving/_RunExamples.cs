using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CSharp.Loading_and_Saving
{
    class _RunExamples
    {
        public static void Main()
        {
            // Run the examples. Un-comment the one you want to run

            CheckFormat.Run();

            // Stop before exiting
            Console.WriteLine("\n\nProgram Finished. Press any key to exit....");
            Console.ReadKey();
        }

        public static String GetDataDir()
        {
            return Path.GetFullPath("../../Loading-and-Saving/Data/");
        }
    }
}
