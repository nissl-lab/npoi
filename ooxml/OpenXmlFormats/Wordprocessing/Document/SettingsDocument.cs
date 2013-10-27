using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Wordprocessing
{

    public class SettingsDocument
    {
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Settings));

        CT_Settings settings = null;
        public SettingsDocument()
        {
            settings = new CT_Settings();
        }
        public static SettingsDocument Parse(Stream stream)
        {
            CT_Settings obj = (CT_Settings)serializer.Deserialize(stream);

            return new SettingsDocument(obj);
        }
        public SettingsDocument(CT_Settings settings)
        {
            this.settings = settings;
        }
        public CT_Settings Settings
        {
            get
            {
                return this.settings;
            }
        }
        public void Save(Stream stream, XmlSerializerNamespaces namespaces)
        {
            serializer.Serialize(stream, settings, namespaces);
        }

    }
}
