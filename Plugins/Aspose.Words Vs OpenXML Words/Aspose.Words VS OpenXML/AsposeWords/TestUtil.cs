// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML
{
    public class TestUtil
    {
        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;

            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

            if (!Directory.Exists(ArtifactsDir))
                Directory.CreateDirectory(ArtifactsDir);
        }

        [SetUp]
        public void SetUp()
        {
            Console.WriteLine($"Clr: {RuntimeInformation.FrameworkDescription}\n");
        }

        [OneTimeTearDown]
        public void OneTimeTearDown()
        {
            ServicePointManager.ServerCertificateValidationCallback = delegate { return false; };

            if (Directory.Exists(ArtifactsDir))
                Directory.Delete(ArtifactsDir, true);
        }

        /// <summary>
        /// Returns the code-base directory.
        /// </summary>
        internal static string GetCodeBaseDir(Assembly assembly)
        {
            // CodeBase is a full URI, such as file:///x:\blahblah.
            Uri uri = new Uri(assembly.Location);
            string mainFolder = Path.GetDirectoryName(uri.LocalPath)
                ?.Substring(0, uri.LocalPath.IndexOf("Aspose.Words VS OpenXML", StringComparison.Ordinal));
            return mainFolder;
        }

        /// <summary>
        /// Returns the assembly directory correctly even if the assembly is shadow-copied.
        /// </summary>
        internal static string GetAssemblyDir(Assembly assembly)
        {
            // CodeBase is a full URI, such as file:///x:\blahblah.
            Uri uri = new Uri(assembly.Location);
            return Path.GetDirectoryName(uri.LocalPath) + Path.DirectorySeparatorChar;
        }

        /// <summary>
        /// Gets the path to the currently running executable.
        /// </summary>
        internal static string AssemblyDir { get; }

        /// <summary>
        /// Gets the path to the codebase directory.
        /// </summary>
        internal static string CodeBaseDir { get; }

        /// <summary>
        /// Gets the path to the documents used by the code examples. Ends with a back slash.
        /// </summary>
        internal static string ArtifactsDir { get; }

        /// <summary>
        /// Gets the path to the documents used by the code examples. Ends with a back slash.
        /// </summary>
        internal static string MyDir { get; }

        static TestUtil()
        {
            AssemblyDir = GetAssemblyDir(Assembly.GetExecutingAssembly());
            CodeBaseDir = GetCodeBaseDir(Assembly.GetExecutingAssembly());
            ArtifactsDir = new Uri(new Uri(CodeBaseDir), @"Data/Artifacts/").LocalPath;
            MyDir = new Uri(new Uri(CodeBaseDir), @"Data/").LocalPath;
        }
    }
}