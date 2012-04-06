using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats
{
    public class PropertiesDocument
    {
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Properties));

        public static PropertiesDocument Parse(Stream stream)
        {
            CT_Properties obj = (CT_Properties)serializer.Deserialize(stream);

            return new PropertiesDocument(obj);
        }

        public PropertiesDocument(CT_Properties prop)
        {
            this._props = prop;
        }
        public PropertiesDocument()
        {
            //_props = new CT_Properties();
        }

        CT_Properties _props;
        public CT_Properties GetProperties()
        {
            return _props;
        }

        public CT_Properties AddNewProperties()
        {
            _props = new CT_Properties();
            return _props;
        }

        public PropertiesDocument Copy()
        {
            PropertiesDocument pd = new PropertiesDocument();
            pd._props = this._props.Copy();
            return pd;
        }

        public void Save(Stream stream,Dictionary<String, String> map)
        {
            serializer.Serialize(stream, _props);
        }
        public override string ToString()
        {
            StringWriter stringWriter = new StringWriter();
            serializer.Serialize(stringWriter, _props);
            return stringWriter.ToString();
        }
    }
}
