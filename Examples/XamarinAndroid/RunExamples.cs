using System;
using System.IO;

using XamarinAndroid.Rendering_Printing;

namespace XamarinAndroid
{
    class RunExamples
    {
        public static string RunIt()
        {
            // Uncomment the one you want to try out

            //// =====================================================
            //// =====================================================
            //// Rendering and Printing
            //// =====================================================
            //// =====================================================


            return RenderShape.Run();
        }
       
        public static string GetOutputFilePath(String inputFilePath)
        {
            string extension = Path.GetExtension(inputFilePath);
            string filename = Path.GetFileNameWithoutExtension(inputFilePath);            

            return Path.GetDirectoryName(inputFilePath) + "/" + filename + "_out" + extension;
        }
    }
}