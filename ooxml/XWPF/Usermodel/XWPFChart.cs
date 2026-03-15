using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.Util;
using System;
using System.IO;
using System.Xml;

namespace NPOI.XWPF.UserModel
{
    public class XWPFChart : POIXMLDocumentPart
    {
        /**
         * Root element of the Chart part
         */
        private CT_ChartSpace chartSpace;

        /**
         * The Chart within that
         */
        private CT_Chart chart;

        // lazy Initialization
        private long checksum;

        /**
         * Construct a chart from a namespace part.
         *
         * @param part the namespace part holding the chart data,
         * the content type must be <code>application/vnd.Openxmlformats-officedocument.Drawingml.chart+xml</code>
         * 
         * @since POI 4.0.0
         */
        protected XWPFChart(PackagePart part) : base(part)
        {
            XmlDocument xml = ConvertStreamToXml(part.GetInputStream());
            chartSpace = ChartSpaceDocument.Parse(xml, POIXMLDocumentPart.NamespaceManager).GetChartSpace();
            chart = chartSpace.chart;
        }


        protected void OnDocumentRead()
        {
            base.OnDocumentRead();
        }

        /// <summary>
        /// Return the underlying CT_ChartSpace bean, the root element of the Chart part.
        /// </summary>
        /// <returns></returns>
        public CT_ChartSpace GetCTChartSpace()
        {
            return chartSpace;
        }

        /// <summary>
        /// Return the underlying CT_Chart bean, within the Chart Space
        /// </summary>
        /// <returns></returns>
        public CT_Chart GetCTChart()
        {
            return chart;
        }


        protected void Commit()
        {
            using(var stream = GetPackagePart().GetOutputStream()) {
                chartSpace.Save(stream);
            }
        }


        public long GetChecksum()
        {
            if(this.checksum == 0)
            {
                Stream is1 = null;
                byte[] data;
                try
                {
                    is1 = GetPackagePart().GetInputStream();
                    data = IOUtils.ToByteArray(is1);
                }
                catch(IOException e)
                {
                    throw new POIXMLException(e);
                }
                finally
                {
                    try
                    {
                        if(is1 != null)
                            is1.Close();
                    }
                    catch(IOException e)
                    {
                        throw new POIXMLException(e);
                    }
                }
                this.checksum = IOUtils.CalculateChecksum(data);
            }
            return this.checksum;
        }


        public bool Equals(Object obj)
        {
            /**
             * In case two objects ARE Equal, but its not the same instance, this
             * implementation will always run through the whole
             * byte-array-comparison before returning true. If this will turn into a
             * performance issue, two possible approaches are available:<br>
             * a) Use the Checksum only and take the risk that two images might have
             * the same CRC32 sum, although they are not the same.<br>
             * b) Use a second (or third) Checksum algorithm to minimise the chance
             * that two images have the same Checksums but are not equal (e.g.
             * CRC32, MD5 and SHA-1 Checksums, Additionally compare the
             * data-byte-array lengths).
             */
            if(obj == this)
            {
                return true;
            }

            if(obj == null)
            {
                return false;
            }

            if(!(obj is XWPFChart))
            {
                return false;
            }
            return false;
        }


        public override int GetHashCode()
        {
            return GetChecksum().GetHashCode();
        }
    }
}
