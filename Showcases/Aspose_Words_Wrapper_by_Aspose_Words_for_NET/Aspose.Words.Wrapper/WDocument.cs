using System;
using System.IO;
using System.Reflection;
using Aspose.Words.Saving;

namespace Aspose.Words.Wrapper
{
    public class WDocument : Document
    {
        internal WDocument()
        {
        }

        internal WDocument(string fileName) : base(fileName)
        {
        }

        internal WDocument(string fileName, LoadOptions loadOptions) : base(fileName, loadOptions)
        {
        }

        internal WDocument(Stream stream) : base(stream)
        {
        }

        internal WDocument(Stream stream, LoadOptions loadOptions) : base(stream, loadOptions)
        {
        }

        /// <summary>
        /// Save document to file.
        /// </summary>
        public void SaveToFile(string fileName)
        {
            Save(fileName);
        }

        /// <summary>
        /// Save document to file (with options).
        /// </summary>
        public void SaveToFileWithOptions(string fileName, Object saveOptions)
        {
            Save(fileName, (SaveOptions)saveOptions);
        }

        /// <summary>
        /// Save document to stream (with options).
        /// </summary>
        public void SaveToStreamWithOptions(Stream stream, Object saveOptions)
        {
            Save(stream, (SaveOptions)saveOptions);
        }

        public Section GetSection(int sectionNumber)
        {
            return Sections[sectionNumber];
        }

        /// <summary>
        /// Instantiate wrapper object (WParagrpah, WRun etc) or ordinary object (Paragraph, Run etc).
        /// </summary>
        /// <param name="className"></param>
        /// <returns></returns>
        public object CreateNode(string className)
        {
            object node = CreateNode(gWrapperAssembly, "Aspose.Words.Wrapper." + className, this) ??
                          CreateNode(gAwAssembly, "Aspose.Words." + className, this);

            if (node == null)
                throw new Exception(string.Format("Invalid node type '{0}'", className));

            return node;
        }

        private static object CreateNode(Assembly assembly, string className, DocumentBase document)
        {
            Type type = assembly.GetType(className);
            if (type == null)
                return null;

            if (!gNodeType.IsAssignableFrom(type))
                return null;

            ConstructorInfo constructor = type.GetConstructor(new Type[] {typeof(DocumentBase)});
            if (constructor == null)
                return null;

            return constructor.Invoke(new object[] {document});
        }

        private static readonly Assembly gWrapperAssembly = typeof (WDocument).Assembly;
        private static readonly Assembly gAwAssembly = typeof(Document).Assembly;
        private static readonly Type gNodeType = typeof (Node);
    }

    public class WDocumentFactory
    {
        /// <summary>
        /// Create a completely empty document.
        /// </summary>
        public WDocument CreateEmpty()
        {
            return new WDocument();
        }

        /// <summary>
        /// Open document from file.
        /// </summary>
        public WDocument OpenFromFile(string fileName)
        {
            return new WDocument(fileName);
        }

        /// <summary>
        /// Open document from file (with options).
        /// </summary>
        public WDocument OpenFromFileWithOptions(string fileName, object options)
        {
            return new WDocument(fileName, (LoadOptions)options);
        }

        /// <summary>
        /// Open document from stream.
        /// </summary>
        public WDocument OpenFromStream(Stream stream)
        {
            return new WDocument(stream);
        }

        /// <summary>
        /// Open document from stream (with options).
        /// </summary>
        public WDocument OpenFromStreamWithOptions(Stream stream, object options)
        {
            return new WDocument(stream, (LoadOptions)options);
        }
    }
}
