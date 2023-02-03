namespace Runner.MAUI
{
    public class Dirs
    {
        internal static string mExternalAppPath { get; }

        /// <summary>
        /// Gets the path to the documents used by the code examples. Ends with a back slash.
        /// </summary>
        internal static string ArtifactsDir { get; }

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
        /// Gets the path to the documents used by the code examples. Ends with a back slash.
        /// </summary>
        internal static string GoldsDir { get; }

        /// <summary>
        /// Gets the URL of the test image.
        /// </summary>
        internal static string ImageUrl { get; }

        static Dirs()
        {
#if __ANDROID__
            mExternalAppPath = Android.App.Application.Context.GetExternalFilesDir(string.Empty).Path;
#elif __IOS__ || __MAC__ || WINDOWS
            mExternalAppPath = "/Users/vderyusev/Aspose/Aspose.Words-for-.NET/ApiExamples/";
#endif
            ArtifactsDir = Path.Combine(mExternalAppPath, "Data/Artifacts/");
            MyDir = Path.Combine(mExternalAppPath, "Data/");
            ImageDir = Path.Combine(mExternalAppPath, "Data/Images/");
            DatabaseDir = Path.Combine(mExternalAppPath, "Data/Database/");
            GoldsDir = Path.Combine(mExternalAppPath, "Data/Golds/");
            FontsDir = Path.Combine(mExternalAppPath, "Data/MyFonts/");
            ImageUrl = new Uri("https://www.aspose.cloud/templates/aspose/img/products/words/aspose_words.svg").AbsoluteUri;
        }
    }
}
