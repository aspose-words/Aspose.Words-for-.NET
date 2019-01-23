// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    /// <summary>
    /// Provides common infrastructure for all API examples that are implemented as unit tests.
    /// </summary>
    public class ApiExampleBase
    {
        [SetUp]
        public void SetUp()
        {
            if (CheckForSkipMono() && IsRunningOnMono())
            {
                Assert.Ignore("Test skipped on mono");
            }

            if (!CheckForSkipSetUp())
            {
                SetUnlimitedLicense();
            }

            if (!Directory.Exists(ArtifactsDir))
                //Create new empty directory
                Directory.CreateDirectory(ArtifactsDir);
        }

        [TearDown]
        public void TearDown()
        {
            if (!CheckForSkipTearDown())
            {
                //Delete all dirs and files from directory
                Directory.Delete(ArtifactsDir, true);
            }
        }

        /// <summary>
        /// Checks when we need to skip precondition before test.
        /// </summary>
        private static bool CheckForSkipSetUp()
        {
            bool skipSetup = TestContext.CurrentContext.Test.Properties["Category"].Contains("SkipSetup");
            return skipSetup;
        }

        /// <summary>
        /// Checks when we need to skip post-condition after test.
        /// </summary>
        private static bool CheckForSkipTearDown()
        {
            bool skipSetup = TestContext.CurrentContext.Test.Properties["Category"].Contains("SkipTearDown");
            return skipSetup;
        }

        /// <summary>
        /// Checks when we need to skip post-condition after test.
        /// </summary>
        private static bool CheckForSkipMono()
        {
            bool skipMono = TestContext.CurrentContext.Test.Properties["Category"].Contains("SkipMono");
            return skipMono;
        }

        /// <summary>
        /// Determine if runtime is Mono.
        /// Workaround for .netcore.
        /// </summary>
        /// <returns>True if being executed in Mono, false otherwise.</returns>
        public static bool IsRunningOnMono() {
            return Type.GetType("Mono.Runtime") != null;
        }

        internal static void SetUnlimitedLicense()
        {
            // This is where the test license is on my development machine.
            string testLicenseFileName = Path.Combine(LicenseDir, "Aspose.Words.lic");

            if (File.Exists(testLicenseFileName))
            {
                // This shows how to use an Aspose.Words license when you have purchased one.
                // You don't have to specify full path as shown here. You can specify just the 
                // file name if you copy the license file into the same folder as your application
                // binaries or you add the license to your project as an embedded resource.
                License license = new License();
                license.SetLicense(testLicenseFileName);
            }
        }

        /// <summary>
        /// Returns the codebase directory.
        /// </summary>
        internal static string GetCodeBaseDir(Assembly assembly)
        {
            // CodeBase is a full URI, such as file:///x:\blahblah.
            Uri uri = new Uri(assembly.CodeBase);
            string mainFolder = Path.GetDirectoryName(uri.LocalPath)
                ?.Substring(0, uri.LocalPath.IndexOf("CSharp", StringComparison.Ordinal));
            return mainFolder;
        }

        /// <summary>
        /// Returns the assembly directory correctly even if the assembly is shadow-copied.
        /// </summary>
        internal static string GetAssemblyDir(Assembly assembly)
        {
            // CodeBase is a full URI, such as file:///x:\blahblah.
            Uri uri = new Uri(assembly.CodeBase);
            return Path.GetDirectoryName(uri.LocalPath) + Path.DirectorySeparatorChar;
        }

        /// <summary>
        /// Gets the path to the currently running executable.
        /// </summary>
        internal static string AssemblyDir
        {
            get { return gAssemblyDir; }
        }

        /// <summary>
        /// Gets the path to the documents used by the code examples. Ends with a back slash.
        /// </summary>
        internal static String ArtifactsDir
        {
            get { return gArtifactsDir; }
        }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static String LicenseDir
        {
            get { return gLicenseDir; }
        }

        /// <summary>
        /// Gets the path to the documents used by the code examples. Ends with a back slash.
        /// </summary>
        internal static String GoldsDir
        {
            get { return gGoldsDir; }
        }

        /// <summary>
        /// Gets the path to the documents used by the code examples. Ends with a back slash.
        /// </summary>
        internal static String MyDir
        {
            get { return gMyDir; }
        }

        /// <summary>
        /// Gets the path to the images used by the code examples. Ends with a back slash.
        /// </summary>
        internal static string ImageDir
        {
            get { return gImageDir; }
        }

        /// <summary>
        /// Gets the path of the demo database. Ends with a back slash.
        /// </summary>
        internal static string DatabaseDir
        {
            get { return gDatabaseDir; }
        }

        static ApiExampleBase()
        {
            gAssemblyDir = GetAssemblyDir(Assembly.GetExecutingAssembly());
            gCodeBaseDir = GetCodeBaseDir(Assembly.GetExecutingAssembly());
            gArtifactsDir = new Uri(new Uri(gCodeBaseDir), @"Data/Artifacts/").LocalPath;
            gLicenseDir = new Uri(new Uri(gCodeBaseDir), @"Data/License/").LocalPath;
            gGoldsDir = new Uri(new Uri(gCodeBaseDir), @"Data/Golds/").LocalPath;
            gMyDir = new Uri(new Uri(gCodeBaseDir), @"Data/").LocalPath;
            gImageDir = new Uri(new Uri(gCodeBaseDir), @"Data/Images/").LocalPath;
            gDatabaseDir = new Uri(new Uri(gCodeBaseDir), @"Data/Database/").LocalPath;
        }

        private static readonly String gAssemblyDir;
        private static readonly String gCodeBaseDir;
        private static readonly String gArtifactsDir;
        private static readonly String gLicenseDir;
        private static readonly String gGoldsDir;
        private static readonly String gMyDir;
        private static readonly String gImageDir;
        private static readonly String gDatabaseDir;
    }
}