using System.IO;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class StyleSheetDocument
    {
        static XmlSerializer serializer = new XmlSerializer(typeof(CT_Stylesheet));
        private CT_Stylesheet stylesheet = null;

        public StyleSheetDocument()
        {
            this.stylesheet = new CT_Stylesheet();
        }

        public StyleSheetDocument(CT_Stylesheet stylesheet)
        {
            this.stylesheet = stylesheet;
        }

        public static StyleSheetDocument Parse(Stream stream)
        {
            CT_Stylesheet obj = (CT_Stylesheet)serializer.Deserialize(stream);
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
            serializer.Serialize(stream, stylesheet);
        }
    }
}
