// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
foreach (Manager manager in GetManagers())
{
    foreach (Contract contract in manager.Contracts)
        yield return contract;
}
