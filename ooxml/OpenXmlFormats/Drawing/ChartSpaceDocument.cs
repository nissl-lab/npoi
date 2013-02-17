using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;
using NPOI.OpenXmlFormats.Dml.Chart;
using System.IO;

namespace NPOI.OpenXmlFormats.Drawing
{
    public class ChartSpaceDocument
    {
        static XmlSerializer serializer = new XmlSerializer(typeof(CT_ChartSpace));
        CT_ChartSpace chartSpace = null;

        public ChartSpaceDocument()
        {
        }
        public ChartSpaceDocument(CT_ChartSpace chartspace)
        {
            this.chartSpace = chartspace;
        }
        public static ChartSpaceDocument Parse(Stream stream)
        {
            CT_ChartSpace obj = (CT_ChartSpace)serializer.Deserialize(stream);
            return new ChartSpaceDocument(obj);
        }
        public CT_ChartSpace GetChartSpace()
        {
            return chartSpace;
        }
        public void SetChartSpace(CT_ChartSpace chartspace)
        {
            this.chartSpace = chartspace;
        }
        public void Save(Stream stream)
        {
            serializer.Serialize(stream, chartSpace);
        }
    }
}
