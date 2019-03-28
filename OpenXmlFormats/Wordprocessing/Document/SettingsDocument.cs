using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Wordprocessing
{

    public class SettingsDocument
    {

        CT_Settings settings = null;
        public SettingsDocument()
        {
            settings = new CT_Settings();
        }
        public static SettingsDocument Parse(XmlDocument doc, XmlNamespaceManager NameSpaceManager)
        {
            CT_Settings obj = CT_Settings.Parse(doc.DocumentElement, NameSpaceManager);
            return new SettingsDocument(obj);
        }

        public void Save(Stream stream)
        {
            using (StreamWriter sw = new StreamWriter(stream))
            {
                settings.Write(sw);
            }
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
    }
}
