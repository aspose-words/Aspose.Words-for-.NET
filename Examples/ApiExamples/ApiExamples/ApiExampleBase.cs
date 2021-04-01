// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
using System.Net.Security;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    /// <summary>
    /// Provides common infrastructure for all API examples that are implemented as unit tests.
    /// </summary>
    public class ApiExampleBase
    {
        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;

            ServicePointManager.ServerCertificateValidationCallback = new
                RemoteCertificateValidationCallback
                (
                    delegate { return true; }
                );

            SetUnlimitedLicense();

            if (!Directory.Exists(ArtifactsDir))
                Directory.CreateDirectory(ArtifactsDir);
        }

        [SetUp]
        public void SetUp()
        {
            if (CheckForSkipMono() && IsRunningOnMono())
            {
                Assert.Ignore("Test skipped on mono");
            }

            Console.WriteLine($"Clr: {RuntimeInformation.FrameworkDescription}\n");
        }

        [OneTimeTearDown]
        public void OneTimeTearDown()
        {
            ServicePointManager.ServerCertificateValidationCallback = new
                RemoteCertificateValidationCallback
                (
                    delegate { return false; }
                );

            if (Directory.Exists(ArtifactsDir))
                Directory.Delete(ArtifactsDir, true);
        }

        /// <summary>
        /// Checks when we need to ignore test on mono.
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
        internal static bool IsRunningOnMono() {
            return Type.GetType("Mono.Runtime") != null;
        }        

        internal static void SetUnlimitedLicense()
        {
            // This is where the test license is on my development machine.
            string testLicenseFileName = Path.Combine(LicenseDir, "Aspose.Total.NET.lic");

            if (File.Exists(testLicenseFileName))
            {
                // This shows how to use an Aspose.Words license when you have purchased one.
                // You don't have to specify full path as shown here. You can specify just the 
                // file name if you copy the license file into the same folder as your application
                // binaries or you add the license to your project as an embedded resource.
                License wordsLicense = new License();
                wordsLicense.SetLicense(testLicenseFileName);

                Aspose.Pdf.License pdfLicense = new Aspose.Pdf.License();
                pdfLicense.SetLicense(testLicenseFileName);

                Aspose.BarCode.License barcodeLicense = new Aspose.BarCode.License();
                barcodeLicense.SetLicense(testLicenseFileName);
            }
        }

        /// <summary>
        /// Returns the code-base directory.
        /// </summary>
        internal static string GetCodeBaseDir(Assembly assembly)
        {
            // CodeBase is a full URI, such as file:///x:\blahblah.
            Uri uri = new Uri(assembly.CodeBase);
            string mainFolder = Path.GetDirectoryName(uri.LocalPath)
                ?.Substring(0, uri.LocalPath.IndexOf("ApiExamples", StringComparison.Ordinal));
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
        internal static string AssemblyDir { get; }

        /// <summary>
        /// Gets the path to the codebase directory.
        /// </summary>
        internal static string CodeBaseDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string LicenseDir { get; }

        /// <summary>
        /// Gets the path to the documents used by the code examples. Ends with a back slash.
        /// </summary>
        internal static string ArtifactsDir { get; }
        
        /// <summary>
        /// Gets the path to the documents used by the code examples. Ends with a back slash.
        /// </summary>
        internal static string GoldsDir { get; }

        /// <summary>
        /// Gets the path to the documents used by the code examples. Ends with a back slash.
        /// </summary>
        internal static string MyDir { get; }

        /// <summary>
        /// Gets the path to the images used by the code examples. Ends with a back slash.
        /// </summary>
        internal static string ImageDir { get; }

        /// <summary>
        /// Gets the path of the demo database. Ends with a back slash.
        /// </summary>
        internal static string DatabaseDir { get; }

        /// <summary>
        /// Gets the path of the free fonts. Ends with a back slash.
        /// </summary>
        internal static string FontsDir { get; }

        /// <summary>
        /// Gets the URL of the Aspose logo.
        /// </summary>
        internal static string AsposeLogoUrl { get; }

        static ApiExampleBase()
        {
            AssemblyDir = GetAssemblyDir(Assembly.GetExecutingAssembly());
            CodeBaseDir = GetCodeBaseDir(Assembly.GetExecutingAssembly());
            ArtifactsDir = new Uri(new Uri(CodeBaseDir), @"Data/Artifacts/").LocalPath;
            LicenseDir = new Uri(new Uri(CodeBaseDir), @"Data/License/").LocalPath;
            GoldsDir = new Uri(new Uri(CodeBaseDir), @"Data/Golds/").LocalPath;
            MyDir = new Uri(new Uri(CodeBaseDir), @"Data/").LocalPath;
            ImageDir = new Uri(new Uri(CodeBaseDir), @"Data/Images/").LocalPath;
            DatabaseDir = new Uri(new Uri(CodeBaseDir), @"Data/Database/").LocalPath;
            FontsDir = new Uri(new Uri(CodeBaseDir), @"Data/MyFonts/").LocalPath;
            AsposeLogoUrl = new Uri("https://www.aspose.cloud/templates/aspose/App_Themes/V3/images/words/header/aspose_words-for-net.png").AbsoluteUri;
        }
    }
}