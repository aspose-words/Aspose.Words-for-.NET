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
    public class ExXamlFixedSaveOptions : ApiExampleBase
    {
        //ExStart
        //ExFor:XamlFixedSaveOptions
        //ExFor:XamlFixedSaveOptions.ResourceSavingCallback
        //ExFor:XamlFixedSaveOptions.ResourcesFolder
        //ExFor:XamlFixedSaveOptions.ResourcesFolderAlias
        //ExFor:XamlFixedSaveOptions.SaveFormat
        //ExSummary:Shows how to print the URIs of linked resources created while converting a document to fixed-form .xaml.
        [Test] //ExSkip
        public void ResourceFolder()
        {
            Document doc = new Document(MyDir + "Rendering.docx");
            ResourceUriPrinter callback = new ResourceUriPrinter();

            // Create a "XamlFixedSaveOptions" object, which we can pass to the document's "Save" method
            // to modify how we save the document to the XAML save format.
            XamlFixedSaveOptions options = new XamlFixedSaveOptions();

            Assert.AreEqual(SaveFormat.XamlFixed, options.SaveFormat);

            // Use the "ResourcesFolder" property to assign a folder in the local file system into which
            // Aspose.Words will save all the document's linked resources, such as images and fonts.
            options.ResourcesFolder = ArtifactsDir + "XamlFixedResourceFolder";

            // Use the "ResourcesFolderAlias" property to use this folder
            // when constructing image URIs instead of the resources folder's name.
            options.ResourcesFolderAlias = ArtifactsDir + "XamlFixedFolderAlias";

            options.ResourceSavingCallback = callback;

            // A folder specified by "ResourcesFolderAlias" will need to contain the resources instead of "ResourcesFolder".
            // We must ensure the folder exists before the callback's streams can put their resources into it.
            Directory.CreateDirectory(options.ResourcesFolderAlias);

            doc.Save(ArtifactsDir + "XamlFixedSaveOptions.ResourceFolder.xaml", options);

            foreach (string resource in callback.Resources)
                Console.WriteLine(resource);
            TestResourceFolder(callback); //ExSkip
        }

        /// <summary>
        /// Counts and prints URIs of resources created during conversion to fixed .xaml.
        /// </summary>
        private class ResourceUriPrinter : IResourceSavingCallback
        {
            public ResourceUriPrinter()
            {
                Resources = new List<string>();
            }

            void IResourceSavingCallback.ResourceSaving(ResourceSavingArgs args)
            {
                Resources.Add($"Resource \"{args.ResourceFileName}\"\n\t{args.ResourceFileUri}");

                // If we specified a resource folder alias, we would also need
                // to redirect each stream to put its resource in the alias folder.
                args.ResourceStream = new FileStream(args.ResourceFileUri, FileMode.Create);
                args.KeepResourceStreamOpen = false;
            }

            public List<string> Resources { get; }
        }
        //ExEnd

        private void TestResourceFolder(ResourceUriPrinter callback)
        {
            Assert.AreEqual(15, callback.Resources.Count);
            foreach (string resource in callback.Resources)
                Assert.True(File.Exists(resource.Split('\t')[1]));
        }
    }
}