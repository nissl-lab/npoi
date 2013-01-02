using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class WorksheetDocument
    {
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Worksheet));
        internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new[] {
            new XmlQualifiedName("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"), 
            new XmlQualifiedName("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships") });

        CT_Worksheet sheet = null;

        public WorksheetDocument()
        {
        }
        public WorksheetDocument(CT_Worksheet sheet)
        {
            this.sheet = sheet;
        }
        public static WorksheetDocument Parse(Stream stream)
        {
            CT_Worksheet obj = (CT_Worksheet)serializer.Deserialize(stream);
            return new WorksheetDocument(obj);
        }
        public CT_Worksheet GetWorksheet()
        {
            return sheet;
        }
        public void SetChartsheet(CT_Worksheet sheet)
        {
            this.sheet = sheet;
        }
        public void Save(Stream stream)
        {
            serializer.Serialize(stream, sheet, namespaces);
        }
    }
}
