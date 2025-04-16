using System;
using System.Collections.Generic;
using DocsExamples.LINQ_Reporting_Engine.Helpers.Data_Source_Objects;

namespace DocsExamples.LINQ_Reporting_Engine.Helpers
{
    internal class Common : DocsExamplesBase
    {
        /// <summary>
        /// Return the first manager from Managers, which is an enumeration of instances of the Manager class. 
        /// </summary>        
        public static Manager GetManager()
        {
            //ExStart:GetManager
            IEnumerator<Manager> managers = GetManagers().GetEnumerator();
            managers.MoveNext();
            
            return managers.Current;
            //ExEnd:GetManager
        }

        /// <summary>
        /// Return an enumeration of instances of the Client class. 
        /// </summary>        
        public static IEnumerable<Client> GetClients()
        {
            //ExStart:GetClients
            foreach (Manager manager in GetManagers())
            {
                foreach (Contract contract in manager.Contracts)
                    yield return contract.Client;
            }
            //ExEnd:GetClients
        }

        /// <summary>
        ///  Return an enumeration of instances of the Manager class.
        /// </summary>
        public static IEnumerable<Manager> GetManagers()
        {
            //ExStart:GetManagers
            Manager manager = new Manager();
            manager.Name = "John Smith";
            manager.Age = 36;
            manager.Photo = Photo();
            Contract initValue = new Contract();
            initValue.Client = new Client();
            initValue.Client.Name = "A Company";
            initValue.Client.Country = "Australia";
            initValue.Client.LocalAddress = "219-241 Cleveland St STRAWBERRY HILLS  NSW  1427";
            initValue.Manager = manager;
            initValue.Price = 1200000;
            initValue.Date = new DateTime(2015, 1, 1);
            Contract initValue2 = new Contract();
            initValue2.Client = new Client();
            initValue2.Client.Name = "B Ltd.";
            initValue2.Client.Country = "Brazil";
            initValue2.Client.LocalAddress = "Avenida João Jorge, 112, ap. 31 Vila Industrial Campinas - SP 13035-680";
            initValue2.Manager = manager;
            initValue2.Price = 750000;
            initValue2.Date = new DateTime(2015, 4, 1);
            Contract initValue3 = new Contract();
            initValue3.Client = new Client();
            initValue3.Client.Name = "C & D";
            initValue3.Client.Country = "Canada";
            initValue3.Client.LocalAddress = "101-3485 RUE DE LA MONTAGNE MONTRÉAL (QUÉBEC) H3G 2A6";
            initValue3.Manager = manager;
            initValue3.Price = 350000;
            initValue3.Date = new DateTime(2015, 7, 1);
            manager.Contracts = new[]
            {
initValue,
initValue2,
initValue3            };
            yield return manager;

            manager = new Manager();
            manager.Name = "Tony Anderson";
            manager.Age = 37;
            manager.Photo = Photo();
            Contract initValue4 = new Contract();
            initValue4.Client = new Client();
            initValue4.Client.Name = "E Corp.";
            initValue4.Client.LocalAddress = "445 Mount Eden Road Mount Eden Auckland 1024";
            initValue4.Manager = manager;
            initValue4.Price = 650000;
            initValue4.Date = new DateTime(2015, 2, 1);
            Contract initValue5 = new Contract();
            initValue5.Client = new Client();
            initValue5.Client.Name = "F & Partners";
            initValue5.Client.LocalAddress = "20 Greens Road Tuahiwi Kaiapoi 7691 ";
            initValue5.Manager = manager;
            initValue5.Price = 550000;
            initValue5.Date = new DateTime(2015, 8, 1);
            manager.Contracts = new[]
            {
initValue4,
initValue5,
            };
            yield return manager;

            manager = new Manager();
            manager.Name = "July James";
            manager.Age = 38;
            manager.Photo = Photo();
            Contract initValue6 = new Contract();
            initValue6.Client = new Client();
            initValue6.Client.Name = "G & Co.";
            initValue6.Client.Country = "Greece";
            initValue6.Client.LocalAddress = "Karkisias 6 GR-111 42  ATHINA GRÉCE";
            initValue6.Manager = manager;
            initValue6.Price = 350000;
            initValue6.Date = new DateTime(2015, 2, 1);
            Contract initValue7 = new Contract();
            initValue7.Client = new Client();
            initValue7.Client.Name = "H Group";
            initValue7.Client.Country = "Hungary";
            initValue7.Client.LocalAddress = "Budapest Fiktív utca 82., IV. em./28.2806";
            initValue7.Manager = manager;
            initValue7.Price = 250000;
            initValue7.Date = new DateTime(2015, 5, 1);
            Contract initValue8 = new Contract();
            initValue8.Client = new Client();
            initValue8.Client.Name = "I & Sons";
            initValue8.Client.LocalAddress = "43 Vogel Street Roslyn Palmerston North 4414";
            initValue8.Manager = manager;
            initValue8.Price = 100000;
            initValue8.Date = new DateTime(2015, 7, 1);
            Contract initValue9 = new Contract();
            initValue9.Client = new Client();
            initValue9.Client.Name = "J Ent.";
            initValue9.Client.Country = "Japan";
            initValue9.Client.LocalAddress = "Hakusan 4-Chōme 3-2 Bunkyō-ku, TŌKYŌ 112-0001 Japan";
            initValue9.Manager = manager;
            initValue9.Price = 100000;
            initValue9.Date = new DateTime(2015, 8, 1);
            manager.Contracts = new[]
            {
initValue6,
initValue7,
initValue8,
initValue9            };
            yield return manager;
            //ExEnd:GetManagers
        }

        /// <summary>
        /// Return an array of photo bytes. 
        /// </summary>
        private static byte[] Photo()
        {
            //ExStart:Photo
            // Load the photo and read all bytes
            byte[] logo = System.IO.File.ReadAllBytes(ImagesDir + "Logo.jpg");
            
            return logo;
            //ExEnd:Photo
        }

        /// <summary>
        ///  Return an enumeration of instances of the Contract class.
        /// </summary>
        public static IEnumerable<Contract> GetContracts()
        {
            //ExStart:GetContracts
            foreach (Manager manager in GetManagers())
            {
                foreach (Contract contract in manager.Contracts)
                    yield return contract;
            }
            //ExEnd:GetContracts
        }
    }
}