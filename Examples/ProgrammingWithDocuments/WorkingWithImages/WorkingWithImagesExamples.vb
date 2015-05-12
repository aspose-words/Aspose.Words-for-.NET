Imports Helpers
Imports NUnit.Framework
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Assert = NUnit.Framework.Assert
Imports Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute

Namespace Examples.ProgrammingWithDocuments.WorkingWithImages
	<TestClass, TestFixture> _
	Public Class WorkingWithImagesExamples
        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub AddImageToEachPage()
	            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithImages/AddImageToEachPage")
	
	            AddImageToEachPageExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub AddWatermark()
	            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithImages/AddWatermark")
	
	            AddWatermarkExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub CompressImages()
	            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithImages/CompressImages")
	
	            CompressImagesExample.Program.Main()
        End Sub

	End Class
End Namespace