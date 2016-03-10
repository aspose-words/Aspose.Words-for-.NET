using System;
using System.IO;
using System.Reflection;
using NUnit.Framework;
using Aspose.Words;

namespace QaTests.Tests
{
    /// <summary>
    /// Base class for all tests.
    /// </summary>
    public class QaTestsBase
    {
        [TestFixtureSetUp]
        public void SetUp()
        {
            SetUnlimitedLicense();
        }

        [TestFixtureTearDown]
        public void TearDown()
        {
            DeleteTempFiles();
        }

        internal static void SetUnlimitedLicense()
        {
            if (File.Exists(TestLicenseFileName))
            {
                // This shows how to use an Aspose.Words license when you have purchased one.
                // You don't have to specify full path as shown here. You can specify just the 
                // file name if you copy the license file into the same folder as your application
                // binaries or you add the license to your project as an embedded resource.
                License license = new License();
                license.SetLicense(TestLicenseFileName);
            }
        }

        internal static void DeleteTempFiles()
        {
            foreach (string file in Directory.GetFiles(MyDir))
                if (file.Contains("_OUT"))
                    File.Delete(file);
        }

        internal static void RemoveLicense()
        {
            License license = new License();
            license.SetLicense("");
        }

        /// <summary>
        /// Returns the assembly directory correctly even if the assembly is shadow-copied.
        /// </summary>
        private static string GetAssemblyDir(Assembly assembly)
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
        /// Gets the path to the documents used by the code tests. Ends with a back slash.
        /// </summary>
        internal static string MyDir
        {
            get { return gMyDir; }
        }

        static QaTestsBase()
        {
            gAssemblyDir = GetAssemblyDir(Assembly.GetExecutingAssembly());
            gMyDir = new Uri(new Uri(gAssemblyDir), @"../../Data/").LocalPath;
        }

        private static readonly string gAssemblyDir;
        private static readonly string gMyDir;

        /// <summary>
        /// This is where the test license is on my development machine.
        /// </summary>
        internal const string TestLicenseFileName = @"X:\awuex\Licenses\Aspose.Words.lic";
    }
}
