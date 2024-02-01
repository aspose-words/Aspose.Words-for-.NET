using System;
using System.Collections.Generic;
using System.Linq;
using ApiExamples.TestData.TestClasses;
using Aspose.Words.ApiExamples.HelperClasses.TestClasses;

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

        public static ShareTestClass[] GetShares()
        {
            return new ShareTestClass[]
            {
                new ShareTestClass("Technology", "Consumer Electronics", "AAPL", 6.602835, -0.0054),
                new ShareTestClass("Technology", "Software - Infrastructure", "MSFT", 5.832072, -0.005),
                new ShareTestClass("Technology", "Software - Infrastructure", "ADBE", 0.562561, -0.0274),
                new ShareTestClass("Technology", "Semiconductors", "NVDA", 1.335994, -0.0074),
                new ShareTestClass("Technology", "Semiconductors", "QCOM", 0.462198, 0.0248),
                new ShareTestClass("Communication Services", "Internet Content & Information", "GOOG", 3.771651, 0.011),
                new ShareTestClass("Communication Services", "Entertainment", "DIS", 0.575768, 0.0102),
                new ShareTestClass("Communication Services", "Entertainment", "WBD", 0.116579, -0.0165),
                new ShareTestClass("Consumer Cyclical", "Internet Retail", "AMZN", 3.011482, 0.044),
                new ShareTestClass("Consumer Cyclical", "Auto Manufactures", "TSLA", 1.816734, -0.0018),
                new ShareTestClass("Consumer Cyclical", "Auto Manufactures", "GM", 0.160205, 0.0026),
                new ShareTestClass("Financial", "Credit Services", "V", 1.1, 0.005)
            };
        }

        public static ShareQuoteTestClass[] GetShareQuotes()
        {
            return new ShareQuoteTestClass[]
            {
                new ShareQuoteTestClass(45131, 15232450, 171.32, 172.50, 170.69, 171.98),
                new ShareQuoteTestClass(45132, 13962990, 172.20, 172.70, 171.40, 171.86),
                new ShareQuoteTestClass(45133, 14902060, 171.86, 171.93, 170.31, 171.35),
                new ShareQuoteTestClass(45134, 16962540, 171.64, 173.10, 171.35, 172.00),
                new ShareQuoteTestClass(45135, 15588280, 171.98, 172.40, 170.00, 171.44)
            };
        }
    }
}