using Aspose.Words;

namespace Runner.MAUI
{
    internal class ConfigureTestsFixture : Dirs, IDisposable
    {
        public ConfigureTestsFixture() 
        {
            SetUnlimitedLicense();

            if (Directory.Exists(ArtifactsDir))
            {
                try
                {
                    Directory.Delete(ArtifactsDir, true);
                    Directory.CreateDirectory(ArtifactsDir);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    throw;
                }

            }
            else
            {
                try
                {
                    Directory.CreateDirectory(ArtifactsDir);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    throw;
                }
            }
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(ArtifactsDir, true);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }

        internal static void SetUnlimitedLicense()
        {
            var testLicenseFileName = Path.Combine(mExternalAppPath, "Data/License/Aspose.Words.NET.lic");

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
    }
}
