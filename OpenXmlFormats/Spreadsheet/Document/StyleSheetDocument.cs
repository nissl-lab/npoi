using System.Diagnostics;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class StyleSheetDocument
    {
        private CT_Stylesheet stylesheet = null;

        public StyleSheetDocument()
        {
            this.stylesheet = new CT_Stylesheet();
        }

        public StyleSheetDocument(CT_Stylesheet stylesheet)
        {
            this.stylesheet = stylesheet;
        }

        public static StyleSheetDocument Parse(XmlDocument xmldoc, XmlNamespaceManager namespaceManager)
        {
            CT_Stylesheet obj = CT_Stylesheet.Parse(xmldoc.DocumentElement,namespaceManager);
            return new StyleSheetDocument(obj);
        }

        public void AddNewStyleSheet()
        {
            this.stylesheet = new CT_Stylesheet();
        }
        public CT_Stylesheet GetStyleSheet()
        {
            return this.stylesheet;
        }
        public void Save(Stream stream)
        {
            using (StreamWriter sw1 = new StreamWriter(stream))
            {
                this.stylesheet.Write(sw1);
            }
        }
    }
}
