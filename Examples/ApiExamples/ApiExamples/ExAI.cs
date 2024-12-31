// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
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
            //ExFor:IAiModelText
            //ExFor:IAiModelText.Summarize(Document, SummarizeOptions)
            //ExFor:IAiModelText.Summarize(Document[], SummarizeOptions)
            //ExFor:SummarizeOptions
            //ExFor:SummarizeOptions.#ctor
            //ExFor:SummarizeOptions.SummaryLength
            //ExFor:SummaryLength
            //ExFor:AiModel
            //ExFor:AiModel.Create(AiModelType)
            //ExFor:AiModel.WithApiKey(String)
            //ExFor:AiModelType
            //ExSummary:Shows how to summarize text using OpenAI and Google models.
            Document firstDoc = new Document(MyDir + "Big document.docx");
            Document secondDoc = new Document(MyDir + "Document.docx");

            string apiKey = Environment.GetEnvironmentVariable("API_KEY");
            // Use OpenAI or Google generative language models.
            IAiModelText model = ((OpenAiModel)AiModel.Create(AiModelType.Gpt4OMini).WithApiKey(apiKey)).WithOrganization("Organization").WithProject("Project");

            Document oneDocumentSummary = model.Summarize(firstDoc, new SummarizeOptions() { SummaryLength = SummaryLength.Short });
            oneDocumentSummary.Save(ArtifactsDir + "AI.AiSummarize.One.docx");

            Document multiDocumentSummary = model.Summarize(new Document[] { firstDoc, secondDoc }, new SummarizeOptions() { SummaryLength = SummaryLength.Long });
            multiDocumentSummary.Save(ArtifactsDir + "AI.AiSummarize.Multi.docx");
            //ExEnd:AiSummarize
        }

        [Test, Ignore("This test should be run manually to manage API requests amount")]
        public void AiTranslate()
        {
            //ExStart:AiTranslate
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:IAiModelText.Translate(Document, AI.Language)
            //ExFor:AI.Language
            //ExSummary:Shows how to translate text using Google models.
            Document doc = new Document(MyDir + "Document.docx");

            string apiKey = Environment.GetEnvironmentVariable("API_KEY");
            // Use Google generative language models.
            IAiModelText model = (IAiModelText)AiModel.Create(AiModelType.Gemini15Flash).WithApiKey(apiKey);

            Document translatedDoc = model.Translate(doc, Language.Arabic);
            translatedDoc.Save(ArtifactsDir + "AI.AiTranslate.docx");
            //ExEnd:AiTranslate
        }

        [Test, Ignore("This test should be run manually to manage API requests amount")]
        public void AiGrammar()
        {
            //ExStart:AiGrammar
            //ReleaseVersion:25.1
            //ExFor:IAiModelText.CheckGrammar(Document, CheckGrammarOptions)
            //ExFor:CheckGrammarOptions
            //ExSummary:Shows how to check the grammar of a document.
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
