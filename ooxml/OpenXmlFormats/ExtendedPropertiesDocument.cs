using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats
{
    public class ExtendedPropertiesDocument
    {
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_ExtendedProperties));

        public static ExtendedPropertiesDocument Parse(Stream stream)
        {
            CT_ExtendedProperties obj = (CT_ExtendedProperties)serializer.Deserialize(stream);

            return new ExtendedPropertiesDocument(obj);
        }

        public ExtendedPropertiesDocument(CT_ExtendedProperties prop)
        {
            this._props = prop;
        }
        public ExtendedPropertiesDocument()
        {
        }

        CT_ExtendedProperties _props;
        public CT_ExtendedProperties GetProperties()
        {
            return _props;
        }

        public CT_ExtendedProperties AddNewProperties()
        {
            _props = new CT_ExtendedProperties();
            return _props;
        }

        public ExtendedPropertiesDocument Copy()
        {
            ExtendedPropertiesDocument pd = new ExtendedPropertiesDocument();
            pd._props = this._props.Copy();
            return pd;
        }

        public void Save(Stream stream,Dictionary<String, String> map)
        {
            serializer.Serialize(stream, _props);
        }
        public override string ToString()
        {
            using (StringWriter stringWriter = new StringWriter())
            {
                serializer.Serialize(stringWriter, _props);
                return stringWriter.ToString();
            }
        }
    }
}
