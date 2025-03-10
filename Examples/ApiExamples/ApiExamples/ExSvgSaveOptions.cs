// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Math;
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

        [Test]
        public void SaveOfficeMath()
        {
            //ExStart:SaveOfficeMath
            //GistId:a775441ecb396eea917a2717cb9e8f8f
            //ExFor:NodeRendererBase.Save(String, SvgSaveOptions)
            //ExFor:NodeRendererBase.Save(Stream, SvgSaveOptions)
            //ExSummary:Shows how to pass save options when rendering office math.
            Document doc = new Document(MyDir + "Office math.docx");

            OfficeMath math = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);

            SvgSaveOptions options = new SvgSaveOptions();
            options.TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs;

            math.GetMathRenderer().Save(ArtifactsDir + "SvgSaveOptions.Output.svg", options);
            
            using (MemoryStream stream = new MemoryStream())
                math.GetMathRenderer().Save(stream, options);
            //ExEnd:SaveOfficeMath
        }

        [Test]
        public void MaxImageResolution()
        {
            //ExStart:MaxImageResolution
            //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
            //ExFor:ShapeBase.SoftEdge
            //ExFor:SoftEdgeFormat.Radius
            //ExFor:SoftEdgeFormat.Remove
            //ExFor:SvgSaveOptions.MaxImageResolution
            //ExSummary:Shows how to set limit for image resolution.
            Document doc = new Document(MyDir + "Rendering.docx");

            SvgSaveOptions saveOptions = new SvgSaveOptions();
            saveOptions.MaxImageResolution = 72;

            doc.Save(ArtifactsDir + "SvgSaveOptions.MaxImageResolution.svg", saveOptions);
            //ExEnd:MaxImageResolution
        }

        [Test]
        public void IdPrefixSvg()
        {
            //ExStart:IdPrefixSvg
            //GistId:f86d49dc0e6781b93e576539a01e6ca2
            //ExFor:SvgSaveOptions.IdPrefix
            //ExSummary:Shows how to add a prefix that is prepended to all generated element IDs (svg).
            Document doc = new Document(MyDir + "Id prefix.docx");

            SvgSaveOptions saveOptions = new SvgSaveOptions();
            saveOptions.IdPrefix = "pfx1_";

            doc.Save(ArtifactsDir + "SvgSaveOptions.IdPrefixSvg.html", saveOptions);
            //ExEnd:IdPrefixSvg
        }

        [Test]
        public void RemoveJavaScriptFromLinksSvg()
        {
            //ExStart:RemoveJavaScriptFromLinksSvg
            //GistId:f86d49dc0e6781b93e576539a01e6ca2
            //ExFor:SvgSaveOptions.RemoveJavaScriptFromLinks
            //ExSummary:Shows how to remove JavaScript from the links (svg).
            Document doc = new Document(MyDir + "JavaScript in HREF.docx");

            SvgSaveOptions saveOptions = new SvgSaveOptions();
            saveOptions.RemoveJavaScriptFromLinks = true;

            doc.Save(ArtifactsDir + "SvgSaveOptions.RemoveJavaScriptFromLinksSvg.html", saveOptions);
            //ExEnd:RemoveJavaScriptFromLinksSvg
        }
    }
}
