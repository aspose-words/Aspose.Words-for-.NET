' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim returnCollection As New List(Of Contract)()
For Each manager As Manager In GetManagers()
    For Each contract As Contract In manager.Contracts
        'yield Return contract
        returnCollection.Add(contract)
    Next
Next
Return returnCollection
