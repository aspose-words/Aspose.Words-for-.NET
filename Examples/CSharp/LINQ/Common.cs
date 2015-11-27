//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSharp.LINQ
{
    class Common
    {

        /// <summary>
        /// Return first manager from Managers which is an enumeration of instances of the Manager class. 
        /// </summary>
        public static Manager GetManager()
        {
            IEnumerator<Manager> managers = GetManagers().GetEnumerator();
            managers.MoveNext();

            return managers.Current;
        }

        /// <summary>
        /// Return an enumeration of instances of the Client class. 
        /// </summary>
        public static IEnumerable<Client> GetClients()
        {
            foreach (Manager manager in GetManagers())
            {
                foreach (Contract contract in manager.Contracts)
                    yield return contract.Client;
            }          
        }

        /// <summary>
        ///  Return an enumeration of instances of the Manager class.
        /// </summary>
        public static IEnumerable<Manager> GetManagers()
        {
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
        }

        /// <summary>
        /// Return an array of photo bytes. 
        /// </summary>
        private static byte[] Photo()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LINQ();

            // Load the photo and read all bytes.
            byte[] imgdata = System.IO.File.ReadAllBytes(dataDir + "photo.png");
            return imgdata;
        }

        /// <summary>
        ///  Return an enumeration of instances of the Contract class.
        /// </summary>
        public static IEnumerable<Contract> GetContracts()
        {
            foreach (Manager manager in GetManagers())
            {
                foreach (Contract contract in manager.Contracts)
                    yield return contract;
            }
        }
    }
}
