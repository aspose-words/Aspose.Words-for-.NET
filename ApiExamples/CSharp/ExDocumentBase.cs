using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Loading;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExDocumentBase : ApiExampleBase
    {
        [Test]
        public void DocumentBaseMisc()
        {
            //ExStart
            //ExFor:DocumentBase.ResourceLoadingCallback
            //ExSummary:Shows how to process inserted resources differently.
            Document doc = new Document();
            doc.ResourceLoadingCallback = new InsertImageByNameHandler();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // We usually insert images as a uri or byte array, but there are many other possibilities with ResourceLoadingCallback
            builder.InsertImage("Google Logo");
            builder.InsertImage("Aspose Logo");

            doc.Save(MyDir + @"\Artifacts\DocumentBase.ResourceLoadingCallback.docx");
            //ExEnd
        }

        private class InsertImageByNameHandler : IResourceLoadingCallback
        {
            public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
            {
                if (args.ResourceType == ResourceType.Image)
                {
                    if (args.OriginalUri == "Google Logo")
                    {
                        var webClient = new WebClient();
                        byte[] imageBytes = webClient.DownloadData("http://www.google.com/images/logos/ps_logo2.png");
                        args.SetData(imageBytes);
                        return ResourceLoadingAction.UserProvided;
                    }

                    if (args.OriginalUri == "Aspose Logo")
                    {
                        var webClient = new WebClient();
                        byte[] imageBytes = webClient.DownloadData("https://www.aspose.com/Images/aspose-logo.jpg");
                        args.SetData(imageBytes);
                        return ResourceLoadingAction.UserProvided;
                    }
                }
                // All other images, documents and CSS stylesheets are handled as before
                return ResourceLoadingAction.Default;
            }
        }

        //ExFor:DocumentBase
        //ExFor:DocumentBase.BackgroundShape
        //ExFor:DocumentBase.ImportNode(Node,System.Boolean)
        //ExFor:DocumentBase.ImportNode(Node,System.Boolean,ImportFormatMode)
        //ExFor:DocumentBase.ImportNode(Node,System.Boolean,ImportFormatMode,INodeCloningListener)
        //ExFor:DocumentBase.ImportNode(Node,System.Boolean,INodeCloningListener)
        //ExFor:DocumentBase.PageColor
        //ExFor:DocumentBase.WarningCallback
    }

}
