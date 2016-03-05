using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.SS.UserModel.Charts;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.XSSF.UserModel.Charts
{
    public class AbstractXSSFChartSeries : IChartSeries
    {
        private String titleValue;
        private CellReference titleRef;
        private TitleType? titleType;
        public void SetTitle(string title)
        {
            titleType = TitleType.String;
            titleValue = title;
        }

        public void SetTitle(CellReference titleReference)
        {
            titleType = TitleType.CellReference;
            titleRef = titleReference;
        }

        public string GetTitleString()
        {
            if (titleType == TitleType.String)
            {
                return titleValue;
            }
            throw new InvalidOperationException("Title type is not String.");
        }

        public CellReference GetTitleCellReference()
        {
            if (titleType== TitleType.CellReference)
            {
                return titleRef;
            }
            throw new InvalidOperationException("Title type is not CellReference.");
        }

        public TitleType? GetTitleType()
        {
            return titleType;
        }
        protected bool IsTitleSet
        {
            get
            {
                return titleType != null;
            }
        }

        protected CT_SerTx GetCTSerTx()
        {
            CT_SerTx tx = new CT_SerTx();
            switch (titleType)
            {
                case TitleType.CellReference:
                    tx.AddNewStrRef().f = titleRef.FormatAsString();
                    return tx;
                case TitleType.String:
                    tx.v =  titleValue;
                    return tx;
                default:
                    throw new InvalidOperationException("Unkown title type: " + titleType);
            }
        }
    }
}
