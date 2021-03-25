using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Threading;
using NUnit.Framework;

namespace OpenXml
{
    public class TestUtil
    {
        static TestUtil()
        {
            MainDataDir = GetCodeBaseDir(Assembly.GetExecutingAssembly());
            MyDir = new Uri(new Uri(MainDataDir), @"Data/").LocalPath;
            ArtifactsDir = new Uri(new Uri(MainDataDir), @"Data/Artifacts/").LocalPath;
        }

        [OneTimeSetUp]
        public static void OneTimeSetUp()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;

            if (!Directory.Exists(ArtifactsDir))
                Directory.CreateDirectory(ArtifactsDir);
        }

        [OneTimeTearDown]
        public static void OneTimeTearDown()
        {
            if (Directory.Exists(ArtifactsDir))
                Directory.Delete(ArtifactsDir, true);
        }

        /// <summary>
        /// Returns the code-base directory.
        /// </summary>
        internal static string GetCodeBaseDir(Assembly assembly)
        {
            Uri uri = new Uri(assembly.CodeBase);
            string mainFolder = Path.GetDirectoryName(uri.LocalPath)
                ?.Substring(0, uri.LocalPath.IndexOf("Aspose.Words VS OpenXML", StringComparison.Ordinal));

            return mainFolder;
        }

        /// <summary>
        /// Gets the path to the codebase directory.
        /// </summary>
        internal static string MainDataDir { get; }

        /// <summary>
        /// Gets the path to the documents used by the code examples.
        /// </summary>
        internal static string MyDir { get; }

        /// <summary>
        /// Gets the path to the artifacts used by the code examples.
        /// </summary>
        internal static string ArtifactsDir { get; }
    }
}