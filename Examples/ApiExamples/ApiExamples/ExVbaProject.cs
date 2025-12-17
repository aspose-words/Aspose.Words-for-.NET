// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
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

            Assert.That(project.Name, Is.EqualTo("Aspose.Project"));

            VbaModuleCollection modules = doc.VbaProject.Modules;

            Assert.That(modules.Count, Is.EqualTo(2));

            Assert.That(modules[0].Name, Is.EqualTo("ThisDocument"));
            Assert.That(modules[0].Type, Is.EqualTo(VbaModuleType.DocumentModule));
            Assert.That(modules[0].SourceCode, Is.Null);

            Assert.That(modules[1].Name, Is.EqualTo("Aspose.Module"));
            Assert.That(modules[1].Type, Is.EqualTo(VbaModuleType.ProceduralModule));
            Assert.That(modules[1].SourceCode, Is.EqualTo("New source code"));
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

            Assert.That(originalVbaProject.Name, Is.EqualTo(copyVbaProject.Name));
            Assert.That(originalVbaProject.CodePage, Is.EqualTo(copyVbaProject.CodePage));
            Assert.That(originalVbaProject.IsSigned, Is.EqualTo(copyVbaProject.IsSigned));
            Assert.That(originalVbaProject.Modules.Count, Is.EqualTo(copyVbaProject.Modules.Count));

            for (int i = 0; i < originalVbaProject.Modules.Count; i++)
            {
                Assert.That(originalVbaProject.Modules[i].Name, Is.EqualTo(copyVbaProject.Modules[i].Name));
                Assert.That(originalVbaProject.Modules[i].Type, Is.EqualTo(copyVbaProject.Modules[i].Type));
                Assert.That(originalVbaProject.Modules[i].SourceCode, Is.EqualTo(copyVbaProject.Modules[i].SourceCode));
            }
        }

        //ExStart
        //ExFor:VbaReference
        //ExFor:VbaReference.Type
        //ExFor:VbaReference.LibId
        //ExFor:VbaReferenceCollection
        //ExFor:VbaReferenceCollection.Item(Int32)
        //ExFor:VbaReferenceCollection.Count
        //ExFor:VbaReferenceCollection.RemoveAt(int)
        //ExFor:VbaReferenceCollection.Remove(VbaReference)
        //ExFor:VbaReferenceType
        //ExFor:VbaProject.References
        //ExSummary:Shows how to get/remove an element from the VBA reference collection.
        [Test]//ExSkip
        public void RemoveVbaReference()
        {
            const string brokenPath = @"X:\broken.dll";
            Document doc = new Document(MyDir + "VBA project.docm");
            
            VbaReferenceCollection references = doc.VbaProject.References;
            Assert.That(references.Count, Is.EqualTo(5 ));
            
            for (int i = references.Count - 1; i >= 0; i--)
            {
                VbaReference reference = doc.VbaProject.References[i];
                string path = GetLibIdPath(reference);
                
                if (path == brokenPath)
                    references.RemoveAt(i);
            }
            Assert.That(references.Count, Is.EqualTo(4 ));

            references.Remove(references[1]);
            Assert.That(references.Count, Is.EqualTo(3 ));

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

        [Test]
        public void IsProtected()
        {
            //ExStart:IsProtected
            //GistId:ac8ba4eb35f3fbb8066b48c999da63b0
            //ExFor:VbaProject.IsProtected
            //ExSummary:Shows whether the VbaProject is password protected.
            Document doc = new Document(MyDir + "Vba protected.docm");
            Assert.That(doc.VbaProject.IsProtected, Is.True);
            //ExEnd:IsProtected
        }
    }
}
