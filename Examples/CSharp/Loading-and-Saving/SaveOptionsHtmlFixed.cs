using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class SaveOptionsHtmlFixed
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();
            UseFontFromTargetMachine(dataDir);
            WriteAllCSSrulesinSingleFile(dataDir);
        }

        static void UseFontFromTargetMachine(string dataDir)
        {
            // ExStart:UseFontFromTargetMachine
            // Load the document from disk.
            Document doc = new Document(dataDir + "Test File (doc).doc");

            HtmlFixedSaveOptions options = new HtmlFixedSaveOptions();
            options.UseTargetMachineFonts = true;

            dataDir = dataDir + "UseFontFromTargetMachine_out.html";
            // Save the document to disk.
            doc.Save(dataDir, options);
            // ExEnd:UseFontFromTargetMachine
            Console.WriteLine("\nFonts from target machine are used in saved HtmlFixed file.\nFile saved at " + dataDir);
        }

        static void WriteAllCSSrulesinSingleFile(string dataDir)
        {
            // ExStart:WriteAllCSSrulesinSingleFile
            // Load the document from disk.
            Document doc = new Document(dataDir + "Test File (doc).doc");

            HtmlFixedSaveOptions options = new HtmlFixedSaveOptions();
            //Setting this property to true restores the old behavior (separate files) for compatibility with legacy code. 
            //Default value is false.
            //All CSS rules are written into single file "styles.css
            options.SaveFontFaceCssSeparately = false;

            dataDir = dataDir + "WriteAllCSSrulesinSingleFile_out.html";
            // Save the document to disk.
            doc.Save(dataDir, options);
            // ExEnd:WriteAllCSSrulesinSingleFile
            Console.WriteLine("\nWrite all CSS rules in single file successfully.\nFile saved at " + dataDir);
        }
    }
}
