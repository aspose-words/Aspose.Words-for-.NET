using Android.App;
using Android.OS;
using Android.Widget;
using System;

namespace XamarinAndroid
{
    [Activity(Label = "XamarinAndroid", MainLauncher = true, Icon = "@drawable/icon")]
    public class MainActivity : Activity
    {
        protected override void OnCreate(Bundle bundle)
        {
            base.OnCreate(bundle);

            // Set our view from the "main" layout resource
            SetContentView(Resource.Layout.Main);

            Button btn = FindViewById<Button>(Resource.Id.button1);
            btn.Click += Btn_Click;

            TextView tv = FindViewById<TextView>(Resource.Id.textView1);
            tv.Text = "Open RunExamples.cs. \nIn RunIt() method uncomment the example that you want to run." +
                "=====================================================";
        }

        private void Btn_Click(object sender, System.EventArgs e)
        {
            TextView tv = FindViewById<TextView>(Resource.Id.textView1);
            try
            {
                tv.Text = RunExamples.RunIt();
            }
            catch (Exception ex)
            {
                tv.Text = ex.ToString();
            }
        }
    }
}

