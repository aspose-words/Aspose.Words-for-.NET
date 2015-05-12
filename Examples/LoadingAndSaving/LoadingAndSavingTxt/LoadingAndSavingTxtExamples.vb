Imports Helpers
Imports NUnit.Framework
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Assert = NUnit.Framework.Assert
Imports Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute

Namespace Examples.LoadingAndSaving.LoadingAndSavingTxt
	<TestClass, TestFixture> _
	Public Class LoadingAndSavingTxtExamples
        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub LoadTxt()
	            TestHelper.SetDataDir("LoadingAndSaving/LoadingAndSavingTxt/LoadTxt")
	
	            LoadTxtExample.Program.Main()
        End Sub

	End Class
End Namespace