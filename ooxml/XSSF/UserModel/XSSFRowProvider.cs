using NPOI.OpenXml4Net.DataVirtualization;
using NPOI.Util;
using System;
using System.Collections.Generic;

namespace NPOI.XSSF.UserModel
{
    public class XSSFRowProvider : IItemsProvider<XSSFRow>
    {
        private static POILogger _logger = POILogFactory.GetLogger(typeof(XSSFRowProvider));

        public int FetchCount()
        {
            throw new NotImplementedException();
        }

        public IList<XSSFRow> FetchRange(int startIndex, int count)
        {
            throw new NotImplementedException();
        }
    }

}