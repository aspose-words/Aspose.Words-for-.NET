using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples
{
    public class DocsExamplesBase
    {
        static DocsExamplesBase()
        {
            MainDataDir = GetCodeBaseDir(Assembly.GetExecutingAssembly());
            ArtifactsDir = new Uri(new Uri(MainDataDir), @"Data/Artifacts/").LocalPath;
            MyDir = new Uri(new Uri(MainDataDir), @"Data/").LocalPath;
            ImagesDir = new Uri(new Uri(MainDataDir), @"Data/Images/").LocalPath;
            LicenseDir = new Uri(new Uri(MainDataDir), @"Data/License/").LocalPath;
            DatabaseDir = new Uri(new Uri(MainDataDir), @"Data/Database/").LocalPath;
            FontsDir = new Uri(new Uri(MainDataDir), @"Data/MyFonts/").LocalPath;
        }

        [OneTimeSetUp]
        public static void OneTimeSetUp()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;

            SetUnlimitedLicense();

            if (!Directory.Exists(ArtifactsDir))
                Directory.CreateDirectory(ArtifactsDir);
        }

        [SetUp]
        public static void SetUp()
        {
            Console.WriteLine($"Clr: {RuntimeInformation.FrameworkDescription}\n");
        }

        [OneTimeTearDown]
        public static void OneTimeTearDown()
        {
            if (Directory.Exists(ArtifactsDir))
                Directory.Delete(ArtifactsDir, true);
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

                Aspose.BarCode.License barcodeLicense = new Aspose.BarCode.License();
                barcodeLicense.SetLicense(testLicenseFileName);

            }
        }

        /// <summary>
        /// Returns the code-base directory.
        /// </summary>
        internal static string GetCodeBaseDir(Assembly assembly)
        {
            Uri uri = new Uri(assembly.Location);
            string mainFolder = Path.GetDirectoryName(uri.LocalPath)
                ?.Substring(0, uri.LocalPath.IndexOf("DocsExamples", StringComparison.Ordinal));
            
            return mainFolder;
        }

        /// <summary>
        /// Gets the path to the codebase directory.
        /// </summary>
        internal static string MainDataDir { get; }

        /// <summary>
        /// Gets the path to the documents used by the code examples.
        /// </summary>
        public static string MyDir { get; }

        /// <summary>
        /// Gets the path to the images used by the code examples.
        /// </summary>
        internal static string ImagesDir { get; }

        /// <summary>
        /// Gets the path of the demo database.
        /// </summary>
        internal static string DatabaseDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string LicenseDir { get; }

        /// <summary>
        /// Gets the path to the artifacts used by the code examples.
        /// </summary>
        internal static string ArtifactsDir { get; }

        /// <summary>
        /// Gets the path of the free fonts. Ends with a back slash.
        /// </summary>
        internal static string FontsDir { get; }
    }
}