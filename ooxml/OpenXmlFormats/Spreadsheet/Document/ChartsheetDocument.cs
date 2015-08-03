using System.IO;
using System.Xml.Serialization;
using System.Xml;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class ChartsheetDocument
    {
       CT_Chartsheet sheet = null;

        public ChartsheetDocument()
        {
        }
        public ChartsheetDocument(CT_Chartsheet sheet)
        {
            this.sheet = sheet;
        }
        public static ChartsheetDocument Parse(XmlDocument xmldoc, XmlNamespaceManager nsmgr)
        {
            CT_Chartsheet obj = CT_Chartsheet.Parse(xmldoc.DocumentElement, nsmgr);
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
            this.sheet.Write(stream);
        }
    }
}
