using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Loading_and_Saving
{
    class WorkWithCHM
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();
            
            LoadCHM(dataDir);
        }

        public static void LoadCHM(string dataDir)
        {
            // ExStart:LoadCHM
            LoadOptions options = new LoadOptions
            {
                Encoding = Encoding.GetEncoding("windows-1251")
            };
            Document doc = new Document(dataDir + "help.chm", options);
            // ExEnd:LoadCHM
        }
    }
}
