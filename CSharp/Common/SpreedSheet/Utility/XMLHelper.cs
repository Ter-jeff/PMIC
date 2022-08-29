#define WPF

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using unvell.Common;

namespace unvell.ReoGrid.Utility
{
    internal static class XMLHelper
    {
        private static Dictionary<Type, XmlSerializer> xmlSerializers;

        public static T LoadXML<T>(Stream s) where T : class
        {
            var type = typeof(T);

#if DEBUG
            var sw = Stopwatch.StartNew();
            try
            {
#endif

                XmlSerializer serializer = null;

                if (xmlSerializers == null) xmlSerializers = new Dictionary<Type, XmlSerializer>();

                if (!xmlSerializers.TryGetValue(type, out serializer))
                    xmlSerializers[type] = serializer = new XmlSerializer(type);

                return serializer.Deserialize(s) as T;

#if DEBUG
            }
            finally
            {
                sw.Stop();
                Logger.Log("xml utility", "type loaded: {0}, {1} ms.", type.Name, sw.ElapsedMilliseconds);
            }
#endif
        }

        public static void SaveXML<T>(Stream s, T obj)
        {
            var type = typeof(T);

#if DEBUG
            var sw = Stopwatch.StartNew();
            try
            {
#endif

                XmlSerializer serializer = null;

                if (xmlSerializers == null) xmlSerializers = new Dictionary<Type, XmlSerializer>();

                if (!xmlSerializers.TryGetValue(type, out serializer))
                    xmlSerializers[type] = serializer = new XmlSerializer(type);

                XmlSerializerNamespaces namespaces = null;

                var xmlnsProp = obj.GetType().GetField("xmlns", BindingFlags.Instance
                                                                | BindingFlags.NonPublic | BindingFlags.Public);

                if (xmlnsProp != null) namespaces = xmlnsProp.GetValue(obj) as XmlSerializerNamespaces;

                var settings = new XmlWriterSettings
                {
                    Encoding = Encoding.UTF8
                };

                using (var writer = XmlWriter.Create(s, settings))
                {
                    serializer.Serialize(writer, obj, namespaces);
                }

#if DEBUG
            }
            finally
            {
                sw.Stop();
                Logger.Log("xml utility", "type saved: {0}, {1} ms.", type.Name, sw.ElapsedMilliseconds);
            }
#endif
        }
    }
}