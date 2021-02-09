// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
    public class ExSvgSaveOptions : ApiExampleBase
    {
        [Test]
        public void SaveLikeImage()
        {
            //ExStart
            //ExFor:SvgSaveOptions.FitToViewPort
            //ExFor:SvgSaveOptions.ShowPageBorder
            //ExFor:SvgSaveOptions.TextOutputMode
            //ExFor:SvgTextOutputMode
            //ExSummary:Shows how to mimic the properties of images when converting a .docx document to .svg.
            Document doc = new Document(MyDir + "Document.docx");

            // Configure the SvgSaveOptions object to save with no page borders or selectable text.
            SvgSaveOptions options = new SvgSaveOptions
            {
                FitToViewPort = true,
                ShowPageBorder = false,
                TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs
            };

            doc.Save(ArtifactsDir + "SvgSaveOptions.SaveLikeImage.svg", options);
            //ExEnd
        }

        //ExStart
        //ExFor:SvgSaveOptions
        //ExFor:SvgSaveOptions.ExportEmbeddedImages
        //ExFor:SvgSaveOptions.ResourceSavingCallback
        //ExFor:SvgSaveOptions.ResourcesFolder
        //ExFor:SvgSaveOptions.ResourcesFolderAlias
        //ExFor:SvgSaveOptions.SaveFormat
        //ExSummary:Shows how to manipulate and print the URIs of linked resources created while converting a document to .svg.
        [Test] //ExSkip
        public void SvgResourceFolder()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            SvgSaveOptions options = new SvgSaveOptions
            {
                SaveFormat = SaveFormat.Svg,
                ExportEmbeddedImages = false,
                ResourcesFolder = ArtifactsDir + "SvgResourceFolder",
                ResourcesFolderAlias = ArtifactsDir + "SvgResourceFolderAlias",
                ShowPageBorder = false,

                ResourceSavingCallback = new ResourceUriPrinter()
            };

            Directory.CreateDirectory(options.ResourcesFolderAlias);

            doc.Save(ArtifactsDir + "SvgSaveOptions.SvgResourceFolder.svg", options);
        }

        /// <summary>
        /// Counts and prints URIs of resources contained by as they are converted to .svg.
        /// </summary>
        private class ResourceUriPrinter : IResourceSavingCallback
        {
            void IResourceSavingCallback.ResourceSaving(ResourceSavingArgs args)
            {
                Console.WriteLine($"Resource #{++mSavedResourceCount} \"{args.ResourceFileName}\"");
                Console.WriteLine("\t" + args.ResourceFileUri);
            }

            private int mSavedResourceCount;
        }
        //ExEnd
    }
}