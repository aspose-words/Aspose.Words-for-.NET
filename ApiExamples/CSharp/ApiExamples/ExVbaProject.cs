// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
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
            //ExSummary:Shows how to create a VbaProject from a scratch for using macros.
            Document doc = new Document();

            // Create a new VBA project
            VbaProject project = new VbaProject();
            project.Name = "Aspose.Project";
            doc.VbaProject = project;

            // Create a new module and specify a macro source code
            VbaModule module = new VbaModule();
            module.Name = "Aspose.Module";
            // VbaModuleType values:
            // procedural module - A collection of subroutines and functions
            // ------
            // document module - A type of VBA project item that specifies a module for embedded macros and programmatic access
            // operations that are associated with a document
            // ------
            // class module - A module that contains the definition for a new object. Each instance of a class creates
            // a new object, and procedures that are defined in the module become properties and methods of the object
            // ------
            // designer module - A VBA module that extends the methods and properties of an ActiveX control that has been
            // registered with the project
            module.Type = VbaModuleType.ProceduralModule;
            module.SourceCode = "New source code";

            // Add module to the VBA project
            doc.VbaProject.Modules.Add(module);

            doc.Save(ArtifactsDir + "Document.CreateVBAMacros.docm");
            //ExEnd

            project = new Document(ArtifactsDir + "Document.CreateVBAMacros.docm").VbaProject;

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
            //ExSummary:Shows how to deep clone VbaProject and VbaModule.
            Document doc = new Document(MyDir + "VBA project.docm");
            Document destDoc = new Document();

            // Clone VbaProject to the document
            VbaProject copyVbaProject = doc.VbaProject.Clone();
            destDoc.VbaProject = copyVbaProject;

            // In destination document we already have "Module1", because it was cloned with VbaProject
            // We will need to remove it before cloning
            VbaModule oldVbaModule = destDoc.VbaProject.Modules["Module1"];
            VbaModule copyVbaModule = doc.VbaProject.Modules["Module1"].Clone();
            destDoc.VbaProject.Modules.Remove(oldVbaModule);
            destDoc.VbaProject.Modules.Add(copyVbaModule);

            destDoc.Save(ArtifactsDir + "Document.CloneVbaProject.docm");
            //ExEnd

            VbaProject originalVbaProject = new Document(ArtifactsDir + "Document.CloneVbaProject.docm").VbaProject;

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
    }
}
