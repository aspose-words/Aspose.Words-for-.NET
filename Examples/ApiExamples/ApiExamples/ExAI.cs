// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using NUnit.Framework;
using Aspose.Words;
using System;
using Aspose.Words.AI;

namespace ApiExamples
{
    [TestFixture]
    public class ExAI : ApiExampleBase
    {
        [Test, Explicit("This test should be run manually to manage API requests amount")]
        public void AiSummarize()
        {
            //ExStart:AiSummarize
            //GistId:366eb64fd56dec3c2eaa40410e594182
            //ExFor:GoogleAiModel
            //ExFor:OpenAiModel
            //ExFor:OpenAiModel.WithOrganization(String)
            //ExFor:OpenAiModel.WithProject(String)
            //ExFor:AiModel
            //ExFor:AiModel.Summarize(Document, SummarizeOptions)
            //ExFor:AiModel.Summarize(Document[], SummarizeOptions)
            //ExFor:AiModel.Create(AiModelType)
            //ExFor:AiModel.WithApiKey(String)
            //ExFor:AiModelType
            //ExFor:SummarizeOptions
            //ExFor:SummarizeOptions.#ctor
            //ExFor:SummarizeOptions.SummaryLength
            //ExFor:SummaryLength
            //ExSummary:Shows how to summarize text using OpenAI and Google models.
            Document firstDoc = new Document(MyDir + "Big document.docx");
            Document secondDoc = new Document(MyDir + "Document.docx");

            string apiKey = Environment.GetEnvironmentVariable("API_KEY");
            // Use OpenAI or Google generative language models.
            AiModel model = ((OpenAiModel)AiModel.Create(AiModelType.Gpt4OMini).WithApiKey(apiKey)).WithOrganization("Organization").WithProject("Project");

            SummarizeOptions options = new SummarizeOptions();

            options.SummaryLength = SummaryLength.Short;
            Document oneDocumentSummary = model.Summarize(firstDoc, options);
            oneDocumentSummary.Save(ArtifactsDir + "AI.AiSummarize.One.docx");

            options.SummaryLength = SummaryLength.Long;
            Document multiDocumentSummary = model.Summarize(new Document[] { firstDoc, secondDoc }, options);
            multiDocumentSummary.Save(ArtifactsDir + "AI.AiSummarize.Multi.docx");
            //ExEnd:AiSummarize
        }

        [Test, Explicit("This test should be run manually to manage API requests amount")]
        public void AiTranslate()
        {
            //ExStart:AiTranslate
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:AiModel.Translate(Document, AI.Language)
            //ExFor:AI.Language
            //ExSummary:Shows how to translate text using Google models.
            Document doc = new Document(MyDir + "Document.docx");

            string apiKey = Environment.GetEnvironmentVariable("API_KEY");
            // Use Google generative language models.
            AiModel model = AiModel.Create(AiModelType.GeminiFlashLatest).WithApiKey(apiKey);

            Document translatedDoc = model.Translate(doc, Language.Arabic);
            translatedDoc.Save(ArtifactsDir + "AI.AiTranslate.docx");
            //ExEnd:AiTranslate
        }

        [Test, Explicit("This test should be run manually to manage API requests amount")]
        public void AiGrammar()
        {
            //ExStart:AiGrammar
            //GistId:f86d49dc0e6781b93e576539a01e6ca2
            //ExFor:AiModel.CheckGrammar(Document, CheckGrammarOptions)
            //ExFor:CheckGrammarOptions
            //ExSummary:Shows how to check the grammar of a document.
            Document doc = new Document(MyDir + "Big document.docx");

            string apiKey = Environment.GetEnvironmentVariable("API_KEY");
            // Use OpenAI generative language models.
            AiModel model = AiModel.Create(AiModelType.Gpt4OMini).WithApiKey(apiKey);

            CheckGrammarOptions grammarOptions = new CheckGrammarOptions();
            grammarOptions.ImproveStylistics = true;

            Document proofedDoc = model.CheckGrammar(doc, grammarOptions);
            proofedDoc.Save(ArtifactsDir + "AI.AiGrammar.docx");
            //ExEnd:AiGrammar
        }

        //ExStart:SelfHostedModel
        //GistId:67c1d01ce69d189983b497fd497a7768
        //ExFor:OpenAiModel
        //ExFor:AiModel.Url
        //ExSummary:Shows how to use self-hosted AI model based on OpenAiModel.
        [Test, Ignore("This test should be run manually when you are configuring your model")] //ExSkip
        public void SelfHostedModel()
        {
            Document doc = new Document(MyDir + "Big document.docx");

            string apiKey = Environment.GetEnvironmentVariable("API_KEY");
            // Use OpenAI generative language models.
            AiModel model = new CustomAiModel().WithApiKey(apiKey);
            model.Url = "https://my.a.com/";

            Document translatedDoc = model.Translate(doc, Language.Russian);
            translatedDoc.Save(ArtifactsDir + "AI.SelfHostedModel.docx");
        }

        /// <summary>
        /// Custom self-hosted AI model.
        /// </summary>
        internal class CustomAiModel : OpenAiModel
        {
            /// <summary>
            /// Gets model name.
            /// </summary>
            protected override string Name
            {
                get { return "my-model-24b"; }
            }
        }
        //ExEnd:SelfHostedModel

        [Test]
        public void ChangeDefaultUrl()
        {
            //ExStart:ChangeDefaultUrl
            //GistId:bd7947d9ad5eb092f532604cb15f593b
            //ExFor:AiModel.Url
            //ExSummary:Shows how to change model default url.
            string apiKey = Environment.GetEnvironmentVariable("API_KEY");
            AiModel model = AiModel.Create(AiModelType.Gpt4OMini).WithApiKey(apiKey);
            // Default value "https://api.openai.com/".
            model.Url = "https://my.a.com/";
            //ExEnd:ChangeDefaultUrl

            Assert.That(model.Url, Is.EqualTo("https://my.a.com/"));
        }

        [Test]
        public void ChangeDefaultTimeout()
        {
            //ExStart:ChangeDefaultTimeout
            //GistId:bd7947d9ad5eb092f532604cb15f593b
            //ExFor:AiModel.Timeout
            //ExSummary:Shows how to change model default timeout.
            string apiKey = Environment.GetEnvironmentVariable("API_KEY");
            AiModel model = AiModel.Create(AiModelType.Gpt4OMini).WithApiKey(apiKey);
            // Default value 100000ms.
            model.Timeout = 250000;
            //ExEnd:ChangeDefaultTimeout

            Assert.That(model.Timeout, Is.EqualTo(250000));
        }

        [Test, Explicit("This test should be run manually to manage API requests amount")]
        public void Gemini()
        {
            //ExStart:Gemini
            //GistId:0da8468118377c4860b28603bc95ffe6
            //ExFor:GoogleAiModel
            //ExFor:GoogleAiModel(String, String)
            //ExSummary:Shows how to use google AI model.
            string apiKey = Environment.GetEnvironmentVariable("API_KEY");
            GoogleAiModel model = new GoogleAiModel("gemini-flash-latest", apiKey);

            Document doc = new Document(MyDir + "Big document.docx");
            SummarizeOptions summarizeOptions = new SummarizeOptions() { SummaryLength = SummaryLength.VeryShort };
            Document summary = model.Summarize(doc, summarizeOptions);
            //ExEnd:Gemini
        }
    }
}
