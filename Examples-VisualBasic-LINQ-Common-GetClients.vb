' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim returnCollection As New List(Of Client)()
For Each manager As Manager In GetManagers()
    For Each contract As Contract In manager.Contracts
        returnCollection.Add(contract.Client)
    Next
Next
Return returnCollection
