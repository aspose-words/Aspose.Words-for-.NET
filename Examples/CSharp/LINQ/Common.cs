using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class Common
    {

        /// <summary>
        /// Return first manager from Managers which is an enumeration of instances of the Manager class. 
        /// </summary>        
        public static Manager GetManager()
        {
            // ExStart:GetManager
            IEnumerator<Manager> managers = GetManagers().GetEnumerator();
            managers.MoveNext();

            return managers.Current;
            // ExEnd:GetManager
        }
        
        /// <summary>
        /// Return an enumeration of instances of the Client class. 
        /// </summary>        
        public static IEnumerable<Client> GetClients()
        {
            // ExStart:GetClients
            foreach (Manager manager in GetManagers())
            {
                foreach (Contract contract in manager.Contracts)
                    yield return contract.Client;
            }
            // ExEnd:GetClients
        }        
        /// <summary>
        ///  Return an enumeration of instances of the Manager class.
        /// </summary>
        
        public static IEnumerable<Manager> GetManagers()
        {
            // ExStart:GetManagers
            Manager manager = new Manager { Name = "John Smith", Age = 36, Photo = Photo() };
            manager.Contracts = new Contract[]
            {
                new Contract { Client = new Client { Name = "A Company", Country= "Australia", LocalAddress = "219-241 Cleveland St STRAWBERRY HILLS  NSW  1427" }, Manager = manager, Price = 1200000, Date = new DateTime(2015, 1, 1) }, 
                new Contract { Client = new Client { Name = "B Ltd.",  Country= "Brazil", LocalAddress = "Avenida João Jorge, 112, ap. 31 Vila Industrial Campinas - SP 13035-680"}, Manager = manager, Price = 750000, Date = new DateTime(2015, 4, 1) }, 
                new Contract { Client = new Client { Name = "C & D", Country= "Canada", LocalAddress = "101-3485 RUE DE LA MONTAGNE MONTRÉAL (QUÉBEC) H3G 2A6" }, Manager = manager, Price = 350000, Date = new DateTime(2015, 7, 1) } 
            };
            yield return manager;

            manager = new Manager { Name = "Tony Anderson", Age = 37, Photo = Photo() };
            manager.Contracts = new Contract[]
            {
                new Contract { Client = new Client { Name = "E Corp.", LocalAddress = "445 Mount Eden Road Mount Eden Auckland 1024" }, Manager = manager, Price = 650000, Date = new DateTime(2015, 2, 1) }, 
                new Contract { Client = new Client { Name = "F & Partners", LocalAddress = "20 Greens Road Tuahiwi Kaiapoi 7691 " }, Manager = manager, Price = 550000, Date = new DateTime(2015, 8, 1) }, 
            };
            yield return manager;

            manager = new Manager { Name = "July James", Age = 38, Photo = Photo() };
            manager.Contracts = new Contract[]
            {
                new Contract { Client = new Client { Name = "G & Co.", Country= "Greece", LocalAddress = "Karkisias 6 GR-111 42  ATHINA GRÉCE" }, Manager = manager, Price = 350000, Date = new DateTime(2015, 2, 1) }, 
                new Contract { Client = new Client { Name = "H Group", Country= "Hungary", LocalAddress = "Budapest Fiktív utca 82., IV. em./28.2806" }, Manager = manager, Price = 250000, Date = new DateTime(2015, 5, 1) }, 
                new Contract { Client = new Client { Name = "I & Sons", LocalAddress ="43 Vogel Street Roslyn Palmerston North 4414" }, Manager = manager, Price = 100000, Date = new DateTime(2015, 7, 1) },
                new Contract { Client = new Client { Name = "J Ent." , Country= "Japan", LocalAddress = "Hakusan 4-Chōme 3-2 Bunkyō-ku, TŌKYŌ 112-0001 Japan"}, Manager = manager, Price = 100000, Date = new DateTime(2015, 8, 1) } 
            };
            yield return manager;
            // ExEnd:GetManagers
        }
        
        /// <summary>
        /// Return an array of photo bytes. 
        /// </summary>
      
        private static byte[] Photo()
        {
            // ExStart:Photo
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LINQ();

            // Load the photo and read all bytes.
            byte[] imgdata = System.IO.File.ReadAllBytes(dataDir + "photo.png");
            return imgdata;
            // ExEnd:Photo
        }
        
        /// <summary>
        ///  Return an enumeration of instances of the Contract class.
        /// </summary        
        public static IEnumerable<Contract> GetContracts()
        {
            // ExStart:GetContracts
            foreach (Manager manager in GetManagers())
            {
                foreach (Contract contract in manager.Contracts)
                    yield return contract;
            }
            // ExEnd:GetContracts
        }

    }
}
