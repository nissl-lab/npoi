using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class WorksheetDocument
    {
        static XmlSerializer serializer = new XmlSerializer(typeof(CT_Worksheet));
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
            serializer.Serialize(stream, sheet);
        }
    }
}
