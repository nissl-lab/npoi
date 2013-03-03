using System.IO;
using System.Xml.Serialization;
using System.Xml;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class ChartsheetDocument
    {
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Chartsheet));
        internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
            new XmlQualifiedName("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main") });
        CT_Chartsheet sheet = null;

        public ChartsheetDocument()
        {
        }
        public ChartsheetDocument(CT_Chartsheet sheet)
        {
            this.sheet = sheet;
        }
        public static ChartsheetDocument Parse(Stream stream)
        {
            CT_Chartsheet obj = (CT_Chartsheet)serializer.Deserialize(stream);
            return new ChartsheetDocument(obj);
        }
        public CT_Chartsheet GetChartsheet()
        {
            return sheet;
        }
        public void SetChartsheet(CT_Chartsheet sheet)
        {
            this.sheet = sheet;
        }
        public void Save(Stream stream)
        {
            serializer.Serialize(stream, sheet, namespaces);
        }
    }
}
