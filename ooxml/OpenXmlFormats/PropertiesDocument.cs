using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;

namespace NPOI.OpenXmlFormats
{
    public class PropertiesDocument
    {
        XmlDocument _xmldoc;

        public static PropertiesDocument Parse(Stream stream)
        {
            //stream.Close();
            PropertiesDocument pd = new PropertiesDocument();
            pd._xmldoc=new XmlDocument();
            pd._xmldoc.Load(stream);
            return pd;
        }

        public PropertiesDocument()
        {
            _props = new CT_Properties();
            _xmldoc = new XmlDocument();
        }

        CT_Properties _props;
        public CT_Properties GetProperties()
        {
            return _props;
        }

        public void AddNewProperties()
        {
            _props = new CT_Properties();

        }

        public PropertiesDocument Copy()
        {
            PropertiesDocument pd = new PropertiesDocument();
            XmlDocument doc2 = (XmlDocument)_xmldoc.Clone();
            pd._xmldoc = doc2;
            return pd;
        }

        public void Save(Stream stream,Dictionary<String, String> map)
        {
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(_xmldoc.NameTable);
            foreach (KeyValuePair<string, string> entry in map)
            {
                nsmgr.AddNamespace(entry.Value, entry.Key);
            }
            _xmldoc.Save(stream);
        }
    }
}
