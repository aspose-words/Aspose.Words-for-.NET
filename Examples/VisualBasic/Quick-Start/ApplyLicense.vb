Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports Aspose.Words

Public Class ApplyLicense
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_QuickStart()

        Dim license As New Aspose.Words.License()
        ' This line attempts to set a license from several locations relative to the executable and Aspose.Words.dll.
        ' You can also use the additional overload to load a license from a stream, this is useful for instance when the 
        ' license is stored as an embedded resource 
        Try
            license.SetLicense("Aspose.Words.lic")
            Console.WriteLine("License set successfully.")

        Catch e As Exception
            ' We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license. 
            Console.WriteLine("There was an error setting the license: " & e.Message)
        End Try


    End Sub
End Class
