using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.OOXML.XSSF.UserModel
{
    public class XSSFDxfStyleProvider : DifferentialStyleProvider
    {
        private IIndexedColorMap colorMap;
        private IBorderFormatting border;
        private IFontFormatting font;
        private ExcelNumberFormat number;
        private IPatternFormatting fill;
        private int stripeSize;

        public XSSFDxfStyleProvider(CT_Dxf dxf, int stripeSize, IIndexedColorMap colorMap)
        {
            this.stripeSize = stripeSize;
            this.colorMap = colorMap;
            if (dxf == null)
            {
                border = null;
                font = null;
                number = null;
                fill = null;
            }
            else
            {
                border = dxf.IsSetBorder() ? new XSSFBorderFormatting(dxf.border) : null;
                font = dxf.IsSetFont() ? new XSSFFontFormatting(dxf.font) : null;
                if (dxf.IsSetNumFmt())
                {
                    CT_NumFmt numFmt = dxf.numFmt;
                    number = new ExcelNumberFormat((int)numFmt.numFmtId, numFmt.formatCode);
                }
                else
                {
                    number = null;
                }
                fill = dxf.IsSetFill() ? new XSSFPatternFormatting(dxf.fill) : null;
            }
        }
        public IBorderFormatting BorderFormatting => border;

        public IFontFormatting FontFormatting => font;

        public IPatternFormatting PatternFormatting => fill;

        public ExcelNumberFormat NumberFormat => number;

        public int StripeSize => stripeSize;
    }
}
