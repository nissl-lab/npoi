using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.OpenXmlFormats.Dml.WordProcessing;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XDDF.UserModel.Chart;
using System;
using System.IO;
using System.Xml;

namespace NPOI.XWPF.UserModel
{
    public class XWPFChart : XDDFChart
    {
        /**
         * this object is used to modify drawing properties
         */
        private CT_Inline ctInline;
        // lazy Initialization
        private long checksum;
        /**
	 * this object is used to write embedded part of chart i.e. xlsx file in docx
	 */
        private Stream sheet;
        /// <summary>
        /// constructor to create a new chart in document
        /// </summary>
        protected XWPFChart():base()
        {
            
        }

        /// <summary>
        /// Construct a chart from a namespace part.
        /// </summary>
        /// <param name="part">the package part holding the chart data</param>
        protected XWPFChart(PackagePart part) : base(part)
        {
            XmlDocument xml = ConvertStreamToXml(part.GetInputStream());
            chartSpace = ChartSpaceDocument.Parse(xml, POIXMLDocumentPart.NamespaceManager).GetChartSpace();
            chart = chartSpace.chart;
        }

        internal void Attach(String chartRelId, XWPFRun run)
        {
            ctInline = run.AddChart(chartRelId);
            ctInline.AddNewExtent();
            SetChartBoundingBox(DEFAULT_WIDTH, DEFAULT_HEIGHT);
        }

        protected void OnDocumentRead()
        {
            base.OnDocumentRead();
        }
        /**
     * method to create relationship with embedded part
     * for example writing xlsx file stream into output stream
     * @param chartSheet
     * @return return relation part which used to write relation in .rels file and get relation id
     * @since POI 4.0
     */
        public RelationPart CreateRelationship(XWPFRelation chartRelation, int index)
        {
            return CreateRelationship(chartRelation, XWPFFactory.GetInstance(), index, false);
        }

        /**
         * protected method which used to initialization of sheet output stream 
         * @param sheet
         * @since POI 4.0
         */
        internal void AddEmbeddedWorkSheet(Stream sheet)
        {
            this.sheet=sheet;
        }

        /**
         * this method is used to write workbook object in embedded part of chart
         * return's true in case of successfully write work book in embedded part or return's false
         * @param wb
         * @return return's true in case of successfully write work book in embedded part or return's false
         * @throws IOException
         * @since POI 4.0
         */
        public bool WriteEmbeddedWorkSheet(IWorkbook wb)
        {
            if(this.sheet!=null && wb!=null)

            {
                wb.Write(this.sheet);
                return true;
            }
            return false;
        }

        /**
         * initialize in line object
         * @param inline
         * @since POI 4.0
         */
        internal void SetInLine(CT_Inline ctInline)
        {
            this.ctInline=ctInline;
        }
        /// <summary>
        /// set chart width
        /// </summary>
        /// <param name="width"></param>
        public long ChartWidth
        {
            get
            {
                return ctInline.extent.cx;
            }
            set
            {
                ctInline.extent.cx = value;
            }
        }
        /// <summary>
        /// Set chart height
        /// </summary>
        /// <param name="height"></param>
        public long ChartHeight
        {
            get {
                return ctInline.extent.cy;
            }
            set
            {
                ctInline.extent.cy = value;
            }
        }
        /// <summary>
        ///  set chart height and width
        /// </summary>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public void SetChartBoundingBox(long width, long height)
        {
            this.ChartWidth = width;
            this.ChartHeight = height;
        }
        public uint ChartTopMargin
        { 
            get
            {
                return ctInline.distT;
            }
            set {
                ctInline.distT = value;
            }
        }
        public uint ChartBottomMargin
        {
            get
            {
                return ctInline.distB;
            }
            set
            {
                ctInline.distB = value;
            }
        }
        public uint ChartLeftMargin
        {
            get
            {
                return ctInline.distL;
            }
            set
            {
                ctInline.distL = value;
            }
        }
        public uint ChartRightMargin
        {
            get
            {
                return ctInline.distR;
            }
            set
            {
                ctInline.distR = value;
            }
        }
        /// <summary>
        /// set chart margin
        /// </summary>
        /// <param name="top"></param>
        /// <param name="right"></param>
        /// <param name="bottom"></param>
        /// <param name="left"></param>
        public void SetChartMargin(uint top, uint right, uint bottom, uint left)
        {
            this.ChartBottomMargin=(bottom);
            this.ChartRightMargin=(right);
            this.ChartLeftMargin=(left);
            this.ChartRightMargin=(right);
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

        protected override POIXMLRelation GetChartRelation()
        {
            return XWPFRelation.CHART;
        }

        protected override POIXMLRelation GetChartWorkbookRelation()
        {
            return XWPFRelation.WORKBOOK;
        }

        protected override POIXMLFactory GetChartFactory()
        {
            return XWPFFactory.GetInstance();
        }
    }
}
