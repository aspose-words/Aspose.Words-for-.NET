using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class WorkingWithRTF
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();
            RecognizeUtf8Text(dataDir);
        }

        public static void RecognizeUtf8Text(string dataDir)
        {
            //ExStart:RecognizeUtf8Text
            RtfLoadOptions loadOptions = new RtfLoadOptions();
            loadOptions.RecognizeUtf8Text = true;

            Document doc = new Document(dataDir + "Utf8Text.rtf", loadOptions);

            dataDir = dataDir + "RecognizeUtf8Text_out.rtf";
            doc.Save(dataDir);
            //ExEnd:RecognizeUtf8Text
            Console.WriteLine("\nUTF8 text has recognized successfully.\nFile saved at " + dataDir);
        }
    }
}
