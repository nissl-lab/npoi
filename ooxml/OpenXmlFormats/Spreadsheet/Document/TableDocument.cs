using NPOI.OpenXml4Net.Util;
using System;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class TableDocument
    {
        CT_Table ctTable = null;

        public TableDocument()
        { 
        }
        public TableDocument(CT_Table table)
        {
            this.ctTable = table;
        }

        public static TableDocument Parse(XmlDocument xmldoc, XmlNamespaceManager namespaceMgr)
        {
            CT_Table obj = CT_Table.Parse(xmldoc.DocumentElement, namespaceMgr);
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
            using (StreamWriter sw = new StreamWriter(stream))
            {
                ctTable.Write(sw);
            }
        }
    }
}
