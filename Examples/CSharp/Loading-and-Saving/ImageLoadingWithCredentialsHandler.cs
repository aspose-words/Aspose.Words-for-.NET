using System;
using Aspose.Words;
using Aspose.Words.Loading;
using System.Net;
namespace Aspose.Words.Examples.CSharp.Loading_and_Saving
{
    //ExStart:ImageLoadingWithCredentialsHandler
    public class ImageLoadingWithCredentialsHandler : IResourceLoadingCallback
    {
        public ImageLoadingWithCredentialsHandler()
        {
            mWebClient = new WebClient();
        }
        public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
        {
            if (args.ResourceType == ResourceType.Image)
            {
                Uri uri = new Uri(args.Uri);

                if (uri.Host == "www.aspose.com")
                    mWebClient.Credentials = new NetworkCredential("User1", "akjdlsfkjs");
                else
                    mWebClient.Credentials = new NetworkCredential("SomeOtherUserID", "wiurlnlvs");

                // Download the bytes from the location referenced by the URI.
                byte[] imageBytes = mWebClient.DownloadData(args.Uri);

                args.SetData(imageBytes);

                return ResourceLoadingAction.UserProvided;
            }
            else
            {
                return ResourceLoadingAction.Default;
            }
        }

        private WebClient mWebClient;
    }
    //ExEnd:ImageLoadingWithCredentialsHandler
}
