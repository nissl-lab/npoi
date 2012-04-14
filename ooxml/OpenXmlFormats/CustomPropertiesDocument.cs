using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats
{
    public class CustomPropertiesDocument
    {
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_CustomProperties));

        public static CustomPropertiesDocument Parse(Stream stream)
        {
            CT_CustomProperties obj = (CT_CustomProperties)serializer.Deserialize(stream);

            return new CustomPropertiesDocument(obj);
        }

        public CustomPropertiesDocument(CT_CustomProperties prop)
        {
            this._props = prop;
        }
        public CustomPropertiesDocument()
        {
            //_props = new CT_Properties();
        }

        CT_CustomProperties _props;
        public CT_CustomProperties GetProperties()
        {
            return _props;
        }

        public CT_CustomProperties AddNewProperties()
        {
            _props = new CT_CustomProperties();
            return _props;
        }

        public CustomPropertiesDocument Copy()
        {
            CustomPropertiesDocument pd = new CustomPropertiesDocument();
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
