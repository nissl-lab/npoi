using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;

namespace NPOI.XSSF.Streaming
{
    public class SXSSFCreationHelper : ICreationHelper
    {
        //TODO: readonly?
        private static POILogger logger = POILogFactory.GetLogger(typeof(SXSSFCreationHelper));

        private SXSSFWorkbook wb;
        private XSSFCreationHelper helper;

        //TODO: @internal
        public SXSSFCreationHelper(SXSSFWorkbook workbook)
        {
            this.helper = new XSSFCreationHelper(workbook.xssfWorkbook);
            this.wb = workbook;
        }

        public IClientAnchor CreateClientAnchor()
        {
            return helper.CreateClientAnchor();
        }

        public IDataFormat CreateDataFormat()
        {
            return helper.CreateDataFormat();
        }

        public IFormulaEvaluator CreateFormulaEvaluator()
        {
            return new SXSSFFormulaEvaluator(wb);
        }

        public IHyperlink CreateHyperlink(HyperlinkType type)
        {
            return helper.CreateHyperlink(type);
        }

        public IRichTextString CreateRichTextString(string text)
        {
            logger.Log(POILogger.INFO, "SXSSF doesn't support Rich Text Strings, any formatting information will be lost");
            return new XSSFRichTextString(text);
        }

        //TODO: missing methods CreateExtendedColor()
    }
}
