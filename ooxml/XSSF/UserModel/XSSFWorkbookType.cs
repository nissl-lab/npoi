using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOI.XSSF.UserModel
{
    public class XSSFWorkbookType
    {

        public static XSSFWorkbookType XLSX = new XSSFWorkbookType(XSSFRelation.WORKBOOK.ContentType, "xlsx");
        public static XSSFWorkbookType XLSM = new XSSFWorkbookType(XSSFRelation.MACROS_WORKBOOK.ContentType, "xlsm");

        private readonly string _contentType;
        private readonly string _extension;

        private XSSFWorkbookType(string contentType, string extension)
        {
            _contentType = contentType;
            _extension = extension;
        }

        public string ContentType
        {
            get { return _contentType; }
        }

        public string Extension
        {
            get { return _extension; }
        }

    }
}
