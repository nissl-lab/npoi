using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace MonthlySalaryReport
{
    class Program
    {
        public static void Main(string[] args)
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet s1=wb.CreateSheet("Monthly Salary Report");
            IRow headerRow = s1.CreateRow(0);
            headerRow.CreateCell(0).SetCellValue("First Name");
            s1.SetColumnWidth(0, 20 * 256);
            headerRow.CreateCell(1).SetCellValue("Last Name");
            s1.SetColumnWidth(1, 20 * 256);
            headerRow.CreateCell(2).SetCellValue("Salary");
            headerRow.CreateCell(3).SetCellValue("Tax Rate");
            headerRow.CreateCell(4).SetCellValue("Tax");
            headerRow.CreateCell(5).SetCellValue("Delivery");

            int row = 1;
            GenerateRow(s1, row++, "Bill", "Zhang", 5000, 9.0/100);
            GenerateRow(s1, row++, "Amy", "Huang", 8000, 11.0/100);
            GenerateRow(s1, row++, "Tomos", "Johnson", 6000, 9.0/100);
            GenerateRow(s1, row++, "Macro", "Jeep", 12000, 15.0/100);
            s1.ForceFormulaRecalculation = true;

            FileStream fs = File.Create("test.xlsx");
            wb.Write(fs);
            fs.Close();
        }

        static void GenerateRow(ISheet sheet1,int rowid,string firstName, string lastName, double salaryAmount, double taxRate)
        {
            IRow row = sheet1.CreateRow(rowid);
            row.CreateCell(0).SetCellValue(firstName);  //A2
            row.CreateCell(1).SetCellValue(lastName);   //B2
            row.CreateCell(2).SetCellValue(salaryAmount);   //C2
            row.CreateCell(3).SetCellValue(taxRate);        //D2
            row.CreateCell(4).SetCellFormula(string.Format("C{0}*D{0}",rowid+1));
            row.CreateCell(5).SetCellFormula(string.Format("C{0}-E{0}", rowid + 1));
        }
    }
}
