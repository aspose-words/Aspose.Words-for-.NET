using Aspose.Words.AI;
using Aspose.Words;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocsExamples.AI_powered_Features
{
    public class Working_with_AI : DocsExamplesBase
    {
        [Test, Explicit("This test should be run manually to manage API requests amount")]
        public void AiSummarize()
        {
            //ExStart:AiSummarize
            //GistId:1e379bedb2b759c1be24c64aad54d13d
            Document firstDoc = new Document(MyDir + "Big document.docx");
            Document secondDoc = new Document(MyDir + "Document.docx");

            string apiKey = Environment.GetEnvironmentVariable("API_KEY");
            // Use OpenAI or Google generative language models.
            IAiModelText model = ((OpenAiModel)AiModel.Create(AiModelType.Gpt4OMini).WithApiKey(apiKey)).WithOrganization("Organization").WithProject("Project");

            SummarizeOptions options = new SummarizeOptions();

            options.SummaryLength = SummaryLength.Short;
            Document oneDocumentSummary = model.Summarize(firstDoc, options);
            oneDocumentSummary.Save(ArtifactsDir + "AI.AiSummarize.One.docx");

            options.SummaryLength = SummaryLength.Long;
            Document multiDocumentSummary = model.Summarize(new Document[] { firstDoc, secondDoc }, options);
            multiDocumentSummary.Save(ArtifactsDir + "AI.AiSummarize.Multi.docx");
            //ExEnd:AiSummarize
        }

        [Test, Ignore("This test should be run manually to manage API requests amount")]
        public void AiTranslate()
        {
            //ExStart:AiTranslate
            //GistId:ea14b3e44c0233eecd663f783a21c4f6
            Document doc = new Document(MyDir + "Document.docx");

            string apiKey = Environment.GetEnvironmentVariable("API_KEY");
            // Use Google generative language models.
            IAiModelText model = (GoogleAiModel)AiModel.Create(AiModelType.Gemini15Flash).WithApiKey(apiKey);

            Document translatedDoc = model.Translate(doc, Language.Arabic);
            translatedDoc.Save(ArtifactsDir + "AI.AiTranslate.docx");
            //ExEnd:AiTranslate
        }

        [Test, Ignore("This test should be run manually to manage API requests amount")]
        public void AiGrammar()
        {
            //ExStart:AiGrammar
            Document doc = new Document(MyDir + "Big document.docx");

            string apiKey = Environment.GetEnvironmentVariable("API_KEY");
            // Use OpenAI generative language models.
            IAiModelText model = (IAiModelText)AiModel.Create(AiModelType.Gpt4OMini).WithApiKey(apiKey);

            CheckGrammarOptions grammarOptions = new CheckGrammarOptions();
            grammarOptions.ImproveStylistics = true;

            Document proofedDoc = model.CheckGrammar(doc, grammarOptions);
            proofedDoc.Save("AI.AiGrammar.docx");
            //ExEnd:AiGrammar
        }
    }
}
