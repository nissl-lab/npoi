using NPOI.OpenXml4Net.OPC;
using System;
using System.IO;

namespace NPOI.XWPF.UserModel
{
    /**
     * Raw picture data, normally attached to a WordprocessingML Drawing.
     * As a rule, embedded xlsx are stored in the /word/embedded/ part of a WordprocessingML package.
     */
    public class XWPFChartData : POIXMLDocumentPart
    {
        /**
         * Create a new XWPFChartData node
         * @since POI 4.0
         */
        protected XWPFChartData() : base()
        {

        }

        /**
         * Construct XWPFChartData from a package part
         *
         * @param part the package part holding the drawing data,
         * 
         * @since POI 4.0
         */
        public XWPFChartData(PackagePart part) : base(part)
        {

        }

        /**
         * Gets the WorkBook data as a byte array.
         * <p>
         * Note, that this call might be expensive since excel data is copied into a temporary byte array.
         * You can grab the chart data directly from the underlying package part as follows:
         * <br>
         * <code>
         * InputStream is = getPackagePart().getInputStream();
         * </code>
         * </p>
         *
         * @return the embedded excel data.
         * @since POI 4.0
         */
        public Stream GetData()
        {
            try
            {
                return GetPackagePart().GetInputStream();
            }
            catch(IOException e)
            {
                throw new POIXMLException(e);
            }
        }

        /**
         * chartData objects store the actual content in the part directly without keeping a
         * copy like all others therefore we need to handle them differently.
         */
        protected internal override void PrepareForCommit()
        {
            // do not clear the part here
        }
    }
}
