// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
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
        //ExSummary:Shows how to print the filenames of linked images created while converting a document to flow-form .xaml.
        [Test] //ExSkip
        public void ImageFolder()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            ImageUriPrinter callback = new ImageUriPrinter(ArtifactsDir + "XamlFlowImageFolderAlias");

            // Create a "XamlFlowSaveOptions" object, which we can pass to the document's "Save" method
            // to modify how we save the document to the XAML save format.
            XamlFlowSaveOptions options = new XamlFlowSaveOptions();

            Assert.AreEqual(SaveFormat.XamlFlow, options.SaveFormat);

            // Use the "ImagesFolder" property to assign a folder in the local file system into which
            // Aspose.Words will save all the document's linked images.
            options.ImagesFolder = ArtifactsDir + "XamlFlowImageFolder";

            // Use the "ImagesFolderAlias" property to use this folder
            // when constructing image URIs instead of the images folder's name.
            options.ImagesFolderAlias = ArtifactsDir + "XamlFlowImageFolderAlias";

            options.ImageSavingCallback = callback;

            // A folder specified by "ImagesFolderAlias" will need to contain the resources instead of "ImagesFolder".
            // We must ensure the folder exists before the callback's streams can put their resources into it.
            Directory.CreateDirectory(options.ImagesFolderAlias);

            doc.Save(ArtifactsDir + "XamlFlowSaveOptions.ImageFolder.xaml", options);

            foreach (string resource in callback.Resources)
                Console.WriteLine($"{callback.ImagesFolderAlias}/{resource}");
            TestImageFolder(callback); //ExSkip
        }

        /// <summary>
        /// Counts and prints filenames of images while their parent document is converted to flow-form .xaml.
        /// </summary>
        private class ImageUriPrinter : IImageSavingCallback
        {
            public ImageUriPrinter(string imagesFolderAlias)
            {
                ImagesFolderAlias = imagesFolderAlias;
                Resources = new List<string>();
            }

            void IImageSavingCallback.ImageSaving(ImageSavingArgs args)
            {
                Resources.Add(args.ImageFileName);

                // If we specified an image folder alias, we would also need
                // to redirect each stream to put its image in the alias folder.
                args.ImageStream = new FileStream($"{ImagesFolderAlias}/{args.ImageFileName}", FileMode.Create);
                args.KeepImageStreamOpen = false;
            }

            public string ImagesFolderAlias { get; }
            public List<string> Resources { get; }
        }
        //ExEnd

        private void TestImageFolder(ImageUriPrinter callback)
        {
            Assert.AreEqual(9, callback.Resources.Count);
            foreach (string resource in callback.Resources)
                Assert.True(File.Exists($"{callback.ImagesFolderAlias}/{resource}"));
        }
    }
}