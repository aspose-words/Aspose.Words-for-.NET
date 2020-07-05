// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
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
        //ExSummary:Shows how to print the URIs of linked resources created during conversion of a document to fixed-form .xaml.
        [Test] //ExSkip
        public void ResourceFolder()
        {
            // Open a document which contains resources
            Document doc = new Document(MyDir + "Rendering.docx");

            ResourceUriPrinter callback = new ResourceUriPrinter();

            XamlFixedSaveOptions options = new XamlFixedSaveOptions
            {
                SaveFormat = SaveFormat.XamlFixed,
                ResourcesFolder = ArtifactsDir + "XamlFixedResourceFolder",
                ResourcesFolderAlias = ArtifactsDir + "XamlFixedFolderAlias",
                ResourceSavingCallback = callback
            };

            // A folder specified by ResourcesFolderAlias will contain the resources instead of ResourcesFolder
            // We must ensure the folder exists before the streams can put their resources into it
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
                // If we set a folder alias in the SaveOptions object, it will be stored here
                Resources.Add($"Resource \"{args.ResourceFileName}\"\n\t{args.ResourceFileUri}");

                // If we specified a ResourcesFolderAlias we will also need to redirect each stream to put its resource in that folder
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