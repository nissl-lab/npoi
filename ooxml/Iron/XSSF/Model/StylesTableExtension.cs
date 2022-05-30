using NPOI.OpenXmlFormats.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.XSSF.Model
{
    public partial class StylesTable
    {
        public StyleSheetDocument GetDoc()
        {
            return doc;
        }
    }
}
