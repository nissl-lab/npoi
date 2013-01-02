using System.IO;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class TableDocument
    {
        static XmlSerializer serializer = new XmlSerializer(typeof(CT_Table));
        CT_Table ctTable = null;

        public TableDocument()
        { 
        }
        public TableDocument(CT_Table table)
        {
            this.ctTable = table;
        }

        public static TableDocument Parse(Stream stream)
        {
            CT_Table obj = (CT_Table)serializer.Deserialize(stream);
            return new TableDocument(obj);
        }

        public CT_Table GetTable()
        {
            return ctTable;
        }

        public void SetTable(CT_Table table)
        {
            this.ctTable = table;
        }

        public void Save(Stream stream)
        {
            serializer.Serialize(stream, ctTable);
        }
    }
}
