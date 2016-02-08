' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim managers As IEnumerator(Of Manager) = GetManagers().GetEnumerator()
managers.MoveNext()
Return managers.Current
