// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Manager manager = new Manager { Name = "John Smith", Age = 36, Photo = Photo() };
manager.Contracts = new Contract[]
{
    new Contract { Client = new Client { Name = "A Company" }, Manager = manager, Price = 1200000, Date = new DateTime(2015, 1, 1) }, 
    new Contract { Client = new Client { Name = "B Ltd." }, Manager = manager, Price = 750000, Date = new DateTime(2015, 4, 1) }, 
    new Contract { Client = new Client { Name = "C & D" }, Manager = manager, Price = 350000, Date = new DateTime(2015, 7, 1) } 
};
yield return manager;

manager = new Manager { Name = "Tony Anderson", Age = 37, Photo = Photo() };
manager.Contracts = new Contract[]
{
    new Contract { Client = new Client { Name = "E Corp." }, Manager = manager, Price = 650000, Date = new DateTime(2015, 2, 1) }, 
    new Contract { Client = new Client { Name = "F & Partners" }, Manager = manager, Price = 550000, Date = new DateTime(2015, 8, 1) }, 
};
yield return manager;

manager = new Manager { Name = "July James", Age = 38, Photo = Photo() };
manager.Contracts = new Contract[]
{
    new Contract { Client = new Client { Name = "G & Co." }, Manager = manager, Price = 350000, Date = new DateTime(2015, 2, 1) }, 
    new Contract { Client = new Client { Name = "H Group" }, Manager = manager, Price = 250000, Date = new DateTime(2015, 5, 1) }, 
    new Contract { Client = new Client { Name = "I & Sons" }, Manager = manager, Price = 100000, Date = new DateTime(2015, 7, 1) },
    new Contract { Client = new Client { Name = "J Ent." }, Manager = manager, Price = 100000, Date = new DateTime(2015, 8, 1) } 
};
yield return manager;
