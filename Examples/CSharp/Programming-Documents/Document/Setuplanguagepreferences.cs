using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class SetupLanguagePreferences
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            AddJapaneseAsEditinglanguages(dataDir);
            SetRussianAsDefaultEditingLanguage(dataDir);
        }

        private static void AddJapaneseAsEditinglanguages(string dataDir)
        {
            // ExStart:AddJapaneseAsEditinglanguages
            // The path to the documents directory.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LanguagePreferences.AddEditingLanguage(EditingLanguage.Japanese);

            Document doc = new Document(dataDir + @"languagepreferences.docx", loadOptions);

            int localeIdFarEast = doc.Styles.DefaultFont.LocaleIdFarEast;
            if (localeIdFarEast == (int)EditingLanguage.Japanese)
                Console.WriteLine("The document either has no any FarEast language set in defaults or it was set to Japanese originally.");
            else
                Console.WriteLine("The document default FarEast language was set to another than Japanese language originally, so it is not overridden.");
            // ExEnd:AddJapaneseAsEditinglanguages
        }

        private static void SetRussianAsDefaultEditingLanguage(string dataDir)
        {
            // ExStart:SetRussianAsDefaultEditingLanguage
            // The path to the documents directory.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LanguagePreferences.DefaultEditingLanguage = EditingLanguage.Russian;

            Document doc = new Document(dataDir + @"languagepreferences.docx", loadOptions);

            int localeId = doc.Styles.DefaultFont.LocaleId;
            if (localeId == (int)EditingLanguage.Russian)
                Console.WriteLine("The document either has no any language set in defaults or it was set to Russian originally.");
            else
                Console.WriteLine("The document default language was set to another than Russian language originally, so it is not overridden.");
            // ExEnd:SetRussianAsDefaultEditingLanguage
        }
    }
}
