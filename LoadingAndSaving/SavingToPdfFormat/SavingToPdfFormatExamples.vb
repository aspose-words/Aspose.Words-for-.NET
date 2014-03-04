Imports Helpers
Imports NUnit.Framework
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Assert = NUnit.Framework.Assert
Imports Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute

Namespace Examples.LoadingAndSaving.SavingToPdfFormat
	<TestClass, TestFixture> _
	Public Class SavingToPdfFormatExamples
        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub ImageToPdf()
	            TestHelper.SetDataDir("LoadingAndSaving/SavingToPdfFormat/ImageToPdf")
	
	            ImageToPdfExample.Program.Main()
        End Sub

	End Class
End Namespace