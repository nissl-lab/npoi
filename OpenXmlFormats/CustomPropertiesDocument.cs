using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats
{
    public class CustomPropertiesDocument
    {
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_CustomProperties));
        internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
            new XmlQualifiedName("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"),
            new XmlQualifiedName("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")
        });

        CT_CustomProperties _props = null;

        public CustomPropertiesDocument(CT_CustomProperties prop)
        {
            this._props = prop;
        }
        public CustomPropertiesDocument()
        {
            //_props = new CT_Properties();
        }

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

        public static CustomPropertiesDocument Parse(Stream stream)
        {
            CT_CustomProperties obj = (CT_CustomProperties)serializer.Deserialize(stream);
            return new CustomPropertiesDocument(obj);
        }

        public void Save(Stream stream)
        {
            serializer.Serialize(stream, _props, namespaces);
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
