// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace AsposeVisualStudioPluginWords.XML
    {
    /// <summary>
    /// Serialization/Deserialization of objects to/from XML
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public static class DynamicObjectSerializer<T> where T : class
        {
        #region Public

        /// <summary>
        /// Reads xml and loads object in document format.
        /// </summary>
        /// <param name="filePath"> Path of xml file </param>
        /// <returns></returns>
        public static T Load (string filePath)
            {
            T serializableObject = XMLDocumentLoader (null, filePath);
            return serializableObject;
            }

        /// <summary>
        /// Loads an object from an XML file , supplying extra data types.
        /// </summary>
        /// <param name="filePath">File path to load.</param>
        /// <param name="extraTypes">Extra data types .</param>
        /// <returns>Object loaded from an XML file.</returns>
        public static T Load (string filePath, System.Type[] extraTypes)
            {
            T serializableObject = XMLDocumentLoader (extraTypes, filePath);
            return serializableObject;
            }

        /// <summary>
        /// Saves an object to an XML file .
        /// </summary>
        /// <param name="serializableObject">Serializable object to be saved to file.</param>
        /// <param name="path">Path of the file to save the object to.</param>
        public static void Save (T serializableObject, string path)
            {
            SaveXMLDocument (serializableObject, null, path);
            }

        /// <summary>
        /// Saves an object to an XML file in .
        /// </summary>
        /// <param name="serializableObject">Serializable object to be saved to file.</param>
        /// <param name="path">Path of the file to save the object to.</param>
        /// <param name="extraTypes">Extra data types to enable serialization of custom types within the object.</param>
        public static void Save (T serializableObject, string path, System.Type[] extraTypes)
            {
            SaveXMLDocument (serializableObject, extraTypes, path);
            }

        #endregion

        #region Private

        private static FileStream CreateFileStream (string path)
            {
            FileStream fileStream = null;

            fileStream = new FileStream (path, FileMode.OpenOrCreate);
            return fileStream;
            }

        private static T XMLDocumentLoader (System.Type[] extraTypes, string path)
            {
            T serializableObject = null;

            using (TextReader textReader = TextFileReader (path))
                {
                XmlSerializer xmlSerializer = XmlSerializerCreation (extraTypes);
                serializableObject = xmlSerializer.Deserialize (textReader) as T;

                }

            return serializableObject;
            }

        private static TextReader TextFileReader (string path)
            {

            TextReader textReader = null;

            textReader = new StreamReader (path);


            return textReader;
            }

        private static TextWriter TextFileWriter (string path)
            {
            TextWriter textWriter = null;


            textWriter = new StreamWriter (path);

            return textWriter;
            }

        private static XmlSerializer XmlSerializerCreation (System.Type[] extraTypes)
            {
            Type ObjectType = typeof (T);

            XmlSerializer xmlSerializer = null;

            if (extraTypes != null)
                xmlSerializer = new XmlSerializer (ObjectType, extraTypes);
            else
                xmlSerializer = new XmlSerializer (ObjectType);

            return xmlSerializer;
            }

        private static void SaveXMLDocument (T serializableObject, System.Type[] extraTypes, string path)
            {
            using (TextWriter textWriter = TextFileWriter (path))
                {
                XmlSerializer xmlSerializer = XmlSerializerCreation (extraTypes);
                xmlSerializer.Serialize (textWriter, serializableObject);
                }
            }

        #endregion
        }
    }
