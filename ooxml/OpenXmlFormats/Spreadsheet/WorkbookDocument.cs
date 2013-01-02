using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class WorkbookDocument
    {
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Workbook));
        internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new[] {
            new XmlQualifiedName("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"), 
            new XmlQualifiedName("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships") });

        CT_Workbook workbook = null;
        public WorkbookDocument()
        {
        }
        public static WorkbookDocument Parse(Stream stream)
        {
            CT_Workbook obj = (CT_Workbook)serializer.Deserialize(stream);
            return new WorkbookDocument(obj);
        }
        public WorkbookDocument(CT_Workbook workbook)
        {
            this.workbook = workbook;
        }
        public CT_Workbook Workbook
        {
            get
            {
                return this.workbook;
            }
        }
        public void Save(Stream stream)
        {
            serializer.Serialize(stream, workbook, namespaces);
        }
    }
}
