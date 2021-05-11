using System;
using System.IO;
using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents
{
    internal class WorkingWithHyphenation : DocsExamplesBase
    {
        [Test]
        public void HyphenateWordsOfLanguages()
        {
            //ExStart:HyphenateWordsOfLanguages
            Document doc = new Document(MyDir + "German text.docx");

            Hyphenation.RegisterDictionary("en-US", MyDir + "hyph_en_US.dic");
            Hyphenation.RegisterDictionary("de-CH", MyDir + "hyph_de_CH.dic");

            doc.Save(ArtifactsDir + "WorkingWithHyphenation.HyphenateWordsOfLanguages.pdf");
            //ExEnd:HyphenateWordsOfLanguages
        }

        [Test]
        public void LoadHyphenationDictionaryForLanguage()
        {
            //ExStart:LoadHyphenationDictionaryForLanguage
            Document doc = new Document(MyDir + "German text.docx");
            
            Stream stream = File.OpenRead(MyDir + "hyph_de_CH.dic");
            Hyphenation.RegisterDictionary("de-CH", stream);

            doc.Save(ArtifactsDir + "WorkingWithHyphenation.LoadHyphenationDictionaryForLanguage.pdf");
            //ExEnd:LoadHyphenationDictionaryForLanguage
        }

        //ExStart:CustomHyphenation
        [Test] //ExSkip
        public void HyphenationCallback()
        {
            try
            {
                // Register hyphenation callback.
                Hyphenation.Callback = new CustomHyphenationCallback();

                Document document = new Document(MyDir + "German text.docx");
                document.Save(ArtifactsDir + "WorkingWithHyphenation.HyphenationCallback.pdf");
            }
            catch (Exception e) when (e.Message.StartsWith("Missing hyphenation dictionary"))
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                Hyphenation.Callback = null;
            }
        }

        public class CustomHyphenationCallback : IHyphenationCallback
        {
            public void RequestDictionary(string language)
            {
                string dictionaryFolder = MyDir;
                string dictionaryFullFileName;
                switch (language)
                {
                    case "en-US":
                        dictionaryFullFileName = Path.Combine(dictionaryFolder, "hyph_en_US.dic");
                        break;
                    case "de-CH":
                        dictionaryFullFileName = Path.Combine(dictionaryFolder, "hyph_de_CH.dic");
                        break;
                    default:
                        throw new Exception($"Missing hyphenation dictionary for {language}.");
                }
                // Register dictionary for requested language.
                Hyphenation.RegisterDictionary(language, dictionaryFullFileName);
            }
        }
        //ExEnd:CustomHyphenation
    }
}