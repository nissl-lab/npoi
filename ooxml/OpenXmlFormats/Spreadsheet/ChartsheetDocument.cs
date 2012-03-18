using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class ChartsheetDocument
    {
        static XmlSerializer serializer = new XmlSerializer(typeof(CT_Chartsheet));
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
            serializer.Serialize(stream, sheet);
        }
    }
}
