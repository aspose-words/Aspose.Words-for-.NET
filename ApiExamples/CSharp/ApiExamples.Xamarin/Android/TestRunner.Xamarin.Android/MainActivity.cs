using Android.App;
using Android.OS;
using Xamarin.Android.NUnitLite;

namespace TestRunner.Xamarin.Android
{
    [Activity(Label = "Xamarin.Android.TestRunner", MainLauncher = true, Icon = "@drawable/icon")]
    public class MainActivity : TestSuiteActivity
    {
        protected override void OnCreate(Bundle bundle)
        {
            AddTest(typeof(ApiExamples.ApiExampleBase).Assembly);

            // Once you called base.OnCreate(), you cannot add more assemblies.
            base.OnCreate(bundle);
        }
    }
}

