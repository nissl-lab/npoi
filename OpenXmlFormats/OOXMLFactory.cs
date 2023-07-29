using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;
using System.IO;

namespace NPOI.OpenXmlFormats
{
    public class OOXMLFactory<T> where T : new()
    {
        XmlSerializer serializerObj = null; 
        public OOXMLFactory()
        {
            serializerObj = new XmlSerializer(typeof(T));
        }

        public T Parse(Stream stream)
        {
            T obj = (T)serializerObj.Deserialize(stream);
            stream.Close();
            return obj;
        }
        public T Create()
        {
            return new T();
        }
    }
}
