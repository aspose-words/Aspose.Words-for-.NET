// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
IEnumerator<Manager> managers = GetManagers().GetEnumerator();
managers.MoveNext();

return managers.Current;
