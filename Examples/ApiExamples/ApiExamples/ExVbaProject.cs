// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Words;
using Aspose.Words.Vba;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    class ExVbaProject : ApiExampleBase
    {
        [Test]
        public void CreateNewVbaProject()
        {
            //ExStart
            //ExFor:VbaProject.#ctor
            //ExFor:VbaProject.Name
            //ExFor:VbaModule.#ctor
            //ExFor:VbaModule.Name
            //ExFor:VbaModule.Type
            //ExFor:VbaModule.SourceCode
            //ExFor:VbaModuleCollection.Add(VbaModule)
            //ExFor:VbaModuleType
            //ExSummary:Shows how to create a VBA project using macros.
            Document doc = new Document();

            // Create a new VBA project.
            VbaProject project = new VbaProject();
            project.Name = "Aspose.Project";
            doc.VbaProject = project;

            // Create a new module and specify a macro source code.
            VbaModule module = new VbaModule();
            module.Name = "Aspose.Module";
            module.Type = VbaModuleType.ProceduralModule;
            module.SourceCode = "New source code";

            // Add the module to the VBA project.
            doc.VbaProject.Modules.Add(module);

            doc.Save(ArtifactsDir + "VbaProject.CreateVBAMacros.docm");
            //ExEnd

            project = new Document(ArtifactsDir + "VbaProject.CreateVBAMacros.docm").VbaProject;

            Assert.AreEqual("Aspose.Project", project.Name);

            VbaModuleCollection modules = doc.VbaProject.Modules;

            Assert.AreEqual(2, modules.Count);

            Assert.AreEqual("ThisDocument", modules[0].Name);
            Assert.AreEqual(VbaModuleType.DocumentModule, modules[0].Type);
            Assert.Null(modules[0].SourceCode);

            Assert.AreEqual("Aspose.Module", modules[1].Name);
            Assert.AreEqual(VbaModuleType.ProceduralModule, modules[1].Type);
            Assert.AreEqual("New source code", modules[1].SourceCode);
        }

        [Test]
        public void CloneVbaProject()
        {
            //ExStart
            //ExFor:VbaProject.Clone
            //ExFor:VbaModule.Clone
            //ExSummary:Shows how to deep clone a VBA project and module.
            Document doc = new Document(MyDir + "VBA project.docm");
            Document destDoc = new Document();

            VbaProject copyVbaProject = doc.VbaProject.Clone();
            destDoc.VbaProject = copyVbaProject;

            // In the destination document, we already have a module named "Module1"
            // because we cloned it along with the project. We will need to remove the module.
            VbaModule oldVbaModule = destDoc.VbaProject.Modules["Module1"];
            VbaModule copyVbaModule = doc.VbaProject.Modules["Module1"].Clone();
            destDoc.VbaProject.Modules.Remove(oldVbaModule);
            destDoc.VbaProject.Modules.Add(copyVbaModule);

            destDoc.Save(ArtifactsDir + "VbaProject.CloneVbaProject.docm");
            //ExEnd

            VbaProject originalVbaProject = new Document(ArtifactsDir + "VbaProject.CloneVbaProject.docm").VbaProject;

            Assert.AreEqual(copyVbaProject.Name, originalVbaProject.Name);
            Assert.AreEqual(copyVbaProject.CodePage, originalVbaProject.CodePage);
            Assert.AreEqual(copyVbaProject.IsSigned, originalVbaProject.IsSigned);
            Assert.AreEqual(copyVbaProject.Modules.Count, originalVbaProject.Modules.Count);

            for (int i = 0; i < originalVbaProject.Modules.Count; i++)
            {
                Assert.AreEqual(copyVbaProject.Modules[i].Name, originalVbaProject.Modules[i].Name);
                Assert.AreEqual(copyVbaProject.Modules[i].Type, originalVbaProject.Modules[i].Type);
                Assert.AreEqual(copyVbaProject.Modules[i].SourceCode, originalVbaProject.Modules[i].SourceCode);
            }
        }

        //ExStart
        //ExFor:VbaReference
        //ExFor:VbaReference.LibId
        //ExFor:VbaReferenceCollection
        //ExFor:VbaReferenceCollection.Count
        //ExFor:VbaReferenceCollection.RemoveAt(int)
        //ExFor:VbaReferenceCollection.Remove(VbaReference)
        //ExFor:VbaReferenceType
        //ExSummary:Shows how to get/remove an element from the VBA reference collection.
        [Test]
        public void RemoveVbaReference()
        {
            const string brokenPath = @"X:\broken.dll";
            Document doc = new Document(MyDir + "VBA project.docm");
            
            VbaReferenceCollection references = doc.VbaProject.References;
            Assert.AreEqual(5 ,references.Count);
            
            for (int i = references.Count - 1; i >= 0; i--)
            {
                VbaReference reference = doc.VbaProject.References[i];
                string path = GetLibIdPath(reference);
                
                if (path == brokenPath)
                    references.RemoveAt(i);
            }
            Assert.AreEqual(4 ,references.Count);
            
            references.Remove(references[1]);
            Assert.AreEqual(3 ,references.Count);
 
            doc.Save(ArtifactsDir + "VbaProject.RemoveVbaReference.docm"); 
        }
 
        /// <summary>
        /// Returns string representing LibId path of a specified reference. 
        /// </summary>
        private static string GetLibIdPath(VbaReference reference)
        {
            switch (reference.Type)
            {
                case VbaReferenceType.Registered:
                case VbaReferenceType.Original:
                case VbaReferenceType.Control:
                    return GetLibIdReferencePath(reference.LibId);
                case VbaReferenceType.Project:
                    return GetLibIdProjectPath(reference.LibId);
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }
 
        /// <summary>
        /// Returns path from a specified identifier of an Automation type library.
        /// </summary>
        private static string GetLibIdReferencePath(string libIdReference)
        {
            if (libIdReference != null)
            {
                string[] refParts = libIdReference.Split('#');
                if (refParts.Length > 3)
                    return refParts[3];
            }
 
            return "";
        }
 
        /// <summary>
        /// Returns path from a specified identifier of an Automation type library.
        /// </summary>
        private static string GetLibIdProjectPath(string libIdProject)
        {
            return libIdProject != null ? libIdProject.Substring(3) : "";
        }
        //ExEnd
    }
}
