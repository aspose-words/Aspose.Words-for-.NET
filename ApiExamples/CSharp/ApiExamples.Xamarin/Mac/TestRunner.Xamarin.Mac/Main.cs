using AppKit;

namespace TestRunner.Xamarin.Mac
{
    static class MainClass
    {
        static void Main(string[] args)
        {
            string[] testArgs = new string[] { typeof(ApiExamples.ApiExampleBase).Assembly.Location, "-noheader", "-labels", "-result:/Users/vyacheslav/falleretic.Aspose.Words-for-.NET/ApiExamples/result.xml" };
			GuiUnit.TestRunner.Main(testArgs);
        }
    }
}
