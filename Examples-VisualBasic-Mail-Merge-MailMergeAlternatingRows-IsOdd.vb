' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
''' <summary>
''' Returns true if the value is odd; false if the value is even.
''' </summary>
Private Shared Function IsOdd(value As Integer) As Boolean
    ' The code is a bit complex, but otherwise automatic conversion to VB does not work.
    Return ((value / 2) * 2).Equals(value)
End Function
