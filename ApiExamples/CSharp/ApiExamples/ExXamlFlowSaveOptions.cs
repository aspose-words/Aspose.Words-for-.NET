// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExXamlFlowSaveOptions : ApiExampleBase
    {
        //ExStart
        //ExFor:XamlFlowSaveOptions
        //ExFor:XamlFlowSaveOptions.#ctor
        //ExFor:XamlFlowSaveOptions.#ctor(SaveFormat)
        //ExFor:XamlFlowSaveOptions.ImageSavingCallback
        //ExFor:XamlFlowSaveOptions.ImagesFolder
        //ExFor:XamlFlowSaveOptions.ImagesFolderAlias
        //ExFor:XamlFlowSaveOptions.SaveFormat
        //ExSummary:Shows how to print the filenames of linked images created during conversion of a document to flow-form .xaml.
        [Test] //ExSkip
        public void XamlFlowImageFolder()
        {
            // Open a document which contains images
            Document doc = new Document(MyDir + "Rendering.doc");

            XamlFlowSaveOptions options = new XamlFlowSaveOptions
            {
                SaveFormat = SaveFormat.XamlFlow,
                ImagesFolder = ArtifactsDir + "XamlFlowImageFolder",
                ImagesFolderAlias = ArtifactsDir + "XamlFlowImageFolderAlias",
                ImageSavingCallback = new ImageUriPrinter(ArtifactsDir + "XamlFlowImageFolderAlias")
            };

            // A folder specified by ImagesFolderAlias will contain the images instead of ImagesFolder
            // We must ensure the folder exists before the streams can put their images into it
            Directory.CreateDirectory(options.ImagesFolderAlias);

            doc.Save(ArtifactsDir + "XamlFlowImageFolder.xaml", options);
        }

        /// <summary>
        /// Counts and prints filenames of images while their parent document is converted to flow-form .xaml
        /// </summary>
        private class ImageUriPrinter : IImageSavingCallback
        {
            public ImageUriPrinter(string imagesFolderAlias)
            {
                mImagesFolderAlias = imagesFolderAlias;
            }

            void IImageSavingCallback.ImageSaving(ImageSavingArgs args)
            {
                Console.WriteLine($"Image #{++mSavedImageCount} \"{args.ImageFileName}\"");

                // If we specified a ImagesFolderAlias we will also need to redirect each stream to put its image in that folder
                args.ImageStream = new FileStream($"{mImagesFolderAlias}/{args.ImageFileName}", FileMode.Create);
                args.KeepImageStreamOpen = false;
            }

            private int mSavedImageCount;
            private readonly string mImagesFolderAlias;
        }
        //ExEnd
    }
}