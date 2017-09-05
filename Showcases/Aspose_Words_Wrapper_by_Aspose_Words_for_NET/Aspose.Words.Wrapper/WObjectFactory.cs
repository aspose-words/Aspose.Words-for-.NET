using System;
using System.Reflection;

namespace Aspose.Words.Wrapper
{
    public class WObjectFactory
    {
        public object CreateObject(string className)
        {
            Type type = gAwAssembly.GetType("Aspose.Words." + className);
            if (type == null)
                throw new Exception(string.Format("Invalid type '{0}'", className));

            ConstructorInfo constructor = type.GetConstructor(new Type[0]);
            if (constructor == null)
                throw new Exception(string.Format("Invalid type '{0}'", className));

            return constructor.Invoke(new object[0]);
        }

        private static readonly Assembly gAwAssembly = typeof(Document).Assembly;
    }
}