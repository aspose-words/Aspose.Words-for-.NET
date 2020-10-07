using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Loading_and_Saving
{
    class WorkingWithVbaReferenceCollection
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();
            //ExStart: RemoveReferenceFromCollectionOfReferences
            Document doc = new Document(dataDir + "VbaProject.docm");

            // Find and remove the reference with some LibId path.
            const string brokenPath = "brokenPath.dll";
            VbaReferenceCollection references = doc.VbaProject.References;
            for (int i = references.Count - 1; i >= 0; i--)
            {
                VbaReference reference = doc.VbaProject.References.ElementAt(i);
                string path = GetLibIdPath(reference);
                if (path == brokenPath)
                    references.RemoveAt(i);
            }

            doc.Save(dataDir + "NoBrokenRef.docm");
            //ExEnd: RemoveReferenceFromCollectionOfReferences
        }
        //ExStart: GetLibIdAndReferencePath
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
        /// <remarks>
        /// Please see details for the syntax at [MS-OVBA], 2.1.1.8 LibidReference. 
        /// </remarks>
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
        /// <remarks>
        /// Please see details for the syntax at [MS-OVBA], 2.1.1.12 ProjectReference. 
        /// </remarks>
        private static string GetLibIdProjectPath(string libIdProject)
        {
            return (libIdProject != null) ? libIdProject.Substring(3) : "";
        }
        //ExEnd: GetLibIdAndReferencePath
    }
}
