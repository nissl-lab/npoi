using System.IO;
using System.Xml.Serialization;
using NPOI.OpenXmlFormats.Dml;
using System.Diagnostics;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class ThemeDocument
    {
        CT_OfficeStyleSheet stylesheet = null;
        public ThemeDocument()
        {
        }
        public ThemeDocument(CT_OfficeStyleSheet stylesheet)
        {
            this.stylesheet = stylesheet;
        }

        public CT_OfficeStyleSheet GetTheme()
        {
            return stylesheet;
        }

        public void Save(Stream stream)
        {
            using (StreamWriter sw = new StreamWriter(stream))
            {
                this.stylesheet.Write(sw);
            }
        }

        public static ThemeDocument Parse(System.Xml.XmlDocument xmldoc, System.Xml.XmlNamespaceManager namespaceManager)
        {
            CT_OfficeStyleSheet obj = CT_OfficeStyleSheet.Parse(xmldoc.DocumentElement, namespaceManager);
            return new ThemeDocument(obj);
        }
    }
}
