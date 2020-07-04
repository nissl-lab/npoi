using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace DownloadXlsx
{
    public partial class DownloadExcel : System.Web.UI.Page
    {
        protected void Button1_Click(object sender, EventArgs e)
        {
            string filename = "test.xlsx";

            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));

            XSSFWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");

            sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
            int x = 1;
            for (int i = 1; i <= 15; i++)
            {
                IRow row = sheet1.CreateRow(i);
                for (int j = 0; j < 15; j++)
                {
                    row.CreateCell(j).SetCellValue(x++);
                }
            }

            using (var f = File.Create(@"c:\test.xlsx"))
            {
                workbook.Write(f);
            }

            Response.WriteFile(@"c:\test.xlsx");
            //http://social.msdn.microsoft.com/Forums/en-US/3a7bdd79-f926-4a5e-bcb0-ef81b6c09dcf/responseoutputstreamwrite-writes-all-but-insetrs-a-char-every-64k?forum=ncl
            //workbook.Write(Response.OutputStream); cannot be used 
            //root cause: Response.OutputStream will insert unnecessary byte into the response bytes.
            Response.Flush();
            Response.End();
        }
    }
}