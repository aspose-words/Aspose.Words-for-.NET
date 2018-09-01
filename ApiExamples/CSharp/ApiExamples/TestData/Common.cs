using System;
using System.Collections.Generic;
using System.Linq;
using ApiExamples.TestData.TestClasses;

namespace ApiExamples.TestData
{
    public static class Common
    {
        public static IEnumerable<ManagerTestClass> GetManagers()
        {
            ManagerTestClass manager = new ManagerTestClass
            {
                Name = "John Smith",
                Age = 36
            };

            manager.Contracts = new[]
            {
                new ContractTestClass
                {
                    Client = new ClientTestClass
                    {
                        Name = "A Company",
                        Country = "Australia",
                        LocalAddress = "219-241 Cleveland St STRAWBERRY HILLS  NSW  1427"
                    },
                    Manager = manager,
                    Price = 1200000,
                    Date = new DateTime(2017, 1, 1)
                },
                new ContractTestClass
                {
                    Client = new ClientTestClass
                    {
                        Name = "B Ltd.",
                        Country = "Brazil",
                        LocalAddress = "Avenida João Jorge, 112, ap. 31 Vila Industrial Campinas - SP 13035-680"
                    },
                    Manager = manager,
                    Price = 750000,
                    Date = new DateTime(2017, 4, 1)
                },
                new ContractTestClass
                {
                    Client = new ClientTestClass
                    {
                        Name = "C & D",
                        Country = "Canada",
                        LocalAddress = "101-3485 RUE DE LA MONTAGNE MONTRÉAL (QUÉBEC) H3G 2A6"
                    },
                    Manager = manager,
                    Price = 350000,
                    Date = new DateTime(2017, 7, 1)
                }
            };

            yield return manager;

            manager = new ManagerTestClass
            {
                Name = "Tony Anderson",
                Age = 37
            };

            manager.Contracts = new[]
            {
                new ContractTestClass
                {
                    Client = new ClientTestClass
                    {
                        Name = "E Corp.",
                        LocalAddress = "445 Mount Eden Road Mount Eden Auckland 1024"
                    },
                    Manager = manager,
                    Price = 650000,
                    Date = new DateTime(2017, 2, 1)
                },
                new ContractTestClass
                {
                    Client = new ClientTestClass
                    {
                        Name = "F & Partners",
                        LocalAddress = "20 Greens Road Tuahiwi Kaiapoi 7691 "
                    },
                    Manager = manager,
                    Price = 550000,
                    Date = new DateTime(2017, 8, 1)
                }
            };

            yield return manager;

            manager = new ManagerTestClass
            {
                Name = "July James",
                Age = 38
            };

            manager.Contracts = new[]
            {
                new ContractTestClass
                {
                    Client = new ClientTestClass
                    {
                        Name = "G & Co.",
                        Country = "Greece",
                        LocalAddress = "Karkisias 6 GR-111 42  ATHINA GRÉCE"
                    },
                    Manager = manager,
                    Price = 350000,
                    Date = new DateTime(2017, 2, 1)
                },
                new ContractTestClass
                {
                    Client = new ClientTestClass
                    {
                        Name = "H Group",
                        Country = "Hungary",
                        LocalAddress = "Budapest Fiktív utca 82., IV. em./28.2806"
                    },
                    Manager = manager,
                    Price = 250000,
                    Date = new DateTime(2017, 5, 1)
                },
                new ContractTestClass
                {
                    Client = new ClientTestClass
                    {
                        Name = "I & Sons",
                        LocalAddress = "43 Vogel Street Roslyn Palmerston North 4414"
                    },
                    Manager = manager,
                    Price = 100000,
                    Date = new DateTime(2017, 7, 1)
                },
                new ContractTestClass
                {
                    Client = new ClientTestClass
                    {
                        Name = "J Ent.",
                        Country = "Japan",
                        LocalAddress = "Hakusan 4-Chōme 3-2 Bunkyō-ku, TŌKYŌ 112-0001 Japan"
                    },
                    Manager = manager,
                    Price = 100000,
                    Date = new DateTime(2017, 8, 1)
                }
            };

            yield return manager;
        }

        public static IEnumerable<ManagerTestClass> GetEmptyManagers()
        {
            return Enumerable.Empty<ManagerTestClass>();
        }

        public static IEnumerable<ClientTestClass> GetClients()
        {
            foreach (ManagerTestClass manager in GetManagers())
            {
                foreach (ContractTestClass contract in manager.Contracts)
                    yield return contract.Client;
            }
        }

        public static IEnumerable<ContractTestClass> GetContracts()
        {
            foreach (ManagerTestClass manager in GetManagers())
            {
                foreach (ContractTestClass contract in manager.Contracts)
                    yield return contract;
            }
        }
    }
}