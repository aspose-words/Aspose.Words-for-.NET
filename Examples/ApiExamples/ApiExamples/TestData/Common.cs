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
            List<ManagerTestClass> result = new List<ManagerTestClass>();
            ManagerTestClass manager = new ManagerTestClass();
            manager.Name = "John Smith";
            manager.Age = 36;
            ContractTestClass initValue = new ContractTestClass();
            initValue.Client = new ClientTestClass();
            initValue.Client.Name = "A Company";
            initValue.Client.Country = "Australia";
            initValue.Client.LocalAddress = "219-241 Cleveland St STRAWBERRY HILLS  NSW  1427";
            initValue.Manager = manager;
            initValue.Price = 1200000;
            initValue.Date = new DateTime(2017, 1, 1);
            ContractTestClass initValue2 = new ContractTestClass();
            initValue2.Client = new ClientTestClass();
            initValue2.Client.Name = "B Ltd.";
            initValue2.Client.Country = "Brazil";
            initValue2.Client.LocalAddress = "Avenida João Jorge, 112, ap. 31 Vila Industrial Campinas - SP 13035-680";
            initValue2.Manager = manager;
            initValue2.Price = 750000;
            initValue2.Date = new DateTime(2017, 4, 1);
            ContractTestClass initValue3 = new ContractTestClass();
            initValue3.Client = new ClientTestClass();
            initValue3.Client.Name = "C & D";
            initValue3.Client.Country = "Canada";
            initValue3.Client.LocalAddress = "101-3485 RUE DE LA MONTAGNE MONTRÉAL (QUÉBEC) H3G 2A6";
            initValue3.Manager = manager;
            initValue3.Price = 350000;
            initValue3.Date = new DateTime(2017, 7, 1);

            manager.Contracts = new[]
            {
                initValue,
                initValue2,
                initValue3
            };

            result.Add(manager);
            manager = new ManagerTestClass();
            manager.Name = "Tony Anderson";
            manager.Age = 37;
            ContractTestClass initValue4 = new ContractTestClass();
            initValue4.Client = new ClientTestClass();
            initValue4.Client.Name = "E Corp.";
            initValue4.Client.LocalAddress = "445 Mount Eden Road Mount Eden Auckland 1024";
            initValue4.Manager = manager;
            initValue4.Price = 650000;
            initValue4.Date = new DateTime(2017, 2, 1);
            ContractTestClass initValue5 = new ContractTestClass();
            initValue5.Client = new ClientTestClass();
            initValue5.Client.Name = "F & Partners";
            initValue5.Client.LocalAddress = "20 Greens Road Tuahiwi Kaiapoi 7691 ";
            initValue5.Manager = manager;
            initValue5.Price = 550000;
            initValue5.Date = new DateTime(2017, 8, 1);

            manager.Contracts = new[]
            {
                initValue4,
                initValue5
            };

            result.Add(manager);
            manager = new ManagerTestClass();
            manager.Name = "July James";
            manager.Age = 38;
            ContractTestClass initValue6 = new ContractTestClass();
            initValue6.Client = new ClientTestClass();
            initValue6.Client.Name = "G & Co.";
            initValue6.Client.Country = "Greece";
            initValue6.Client.LocalAddress = "Karkisias 6 GR-111 42  ATHINA GRÉCE";
            initValue6.Manager = manager;
            initValue6.Price = 350000;
            initValue6.Date = new DateTime(2017, 2, 1);
            ContractTestClass initValue7 = new ContractTestClass();
            initValue7.Client = new ClientTestClass();
            initValue7.Client.Name = "H Group";
            initValue7.Client.Country = "Hungary";
            initValue7.Client.LocalAddress = "Budapest Fiktív utca 82., IV. em./28.2806";
            initValue7.Manager = manager;
            initValue7.Price = 250000;
            initValue7.Date = new DateTime(2017, 5, 1);
            ContractTestClass initValue8 = new ContractTestClass();
            initValue8.Client = new ClientTestClass();
            initValue8.Client.Name = "I & Sons";
            initValue8.Client.LocalAddress = "43 Vogel Street Roslyn Palmerston North 4414";
            initValue8.Manager = manager;
            initValue8.Price = 100000;
            initValue8.Date = new DateTime(2017, 7, 1);
            ContractTestClass initValue9 = new ContractTestClass();
            initValue9.Client = new ClientTestClass();
            initValue9.Client.Name = "J Ent.";
            initValue9.Client.Country = "Japan";
            initValue9.Client.LocalAddress = "Hakusan 4-Chōme 3-2 Bunkyō-ku, TŌKYŌ 112-0001 Japan";
            initValue9.Manager = manager;
            initValue9.Price = 100000;
            initValue9.Date = new DateTime(2017, 8, 1);

            manager.Contracts = new[]
            {
                initValue6,
                initValue7,
                initValue8,
                initValue9
            };

            result.Add(manager);
            return result;
        }

        public static IEnumerable<ManagerTestClass> GetEmptyManagers()
        {
            return Enumerable.Empty<ManagerTestClass>();
        }

        public static IEnumerable<ClientTestClass> GetClients()
        {
            List<ClientTestClass> result = new List<ClientTestClass>();
            foreach (ManagerTestClass manager in GetManagers())
            {
                foreach (ContractTestClass contract in manager.Contracts)
                    result.Add(contract.Client);
            }

            return result;
        }

        public static IEnumerable<ContractTestClass> GetContracts()
        {
            List<ContractTestClass> result = new List<ContractTestClass>();
            foreach (ManagerTestClass manager in GetManagers())
            {
                foreach (ContractTestClass contract in manager.Contracts)
                    result.Add(contract);
            }

            return result;
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