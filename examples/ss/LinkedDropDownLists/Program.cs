/**
 * Demonstrates one technique that may be used to create linked or dependent
 * drop down lists. This refers to a situation in which the selection made
 * in one drop down list affects the options that are displayed in the second
 * or subsequent drop down list(s). In this example, the value the user selects
 * from the down list in cell A1 will affect the values displayed in the linked
 * drop down list in cell B1. For the sake of simplicity, the data for the drop
 * down lists is included on the same worksheet but this does not have to be the
 * case; the data could appear on a separate sheet. If this were done, then the
 * names for the regions would have to be different, they would have to include
 * the name of the sheet.
 * 
 * There are two keys to this technique. The first is the use of named area or 
 * regions of cells to hold the data for the drop down lists and the second is
 * making use of the INDIRECT() function to convert a name into the addresses
 * of the cells it refers to.
 * 
 * Note that whilst this class builds just two linked drop down lists, there is
 * nothing to prevent more being created. Quite simply, use the value selected
 * by the user in one drop down list to determine what is shown in another and the
 * value selected in that drop down list to determine what is shown in a third,
 * and so on. Also, note that the data for the drop down lists is contained on
 * contained on the same sheet as the validations themselves. This is done simply
 * for simplicity and there is nothing to prevent a separate sheet being created
 * and used to hold the data. If this is done then problems may be encountered
 * if the sheet is opened with OpenOffice Calc. To prevent these problems, it is
 * better to include the name of the sheet when calling the setRefersToFormula()
 * method.
 *
 * @author Mark Beardsley [msb at apache.org]
 * @version 1.00 30th March 2012
 */

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.IO;

namespace LinkedDropDownLists
{
    class Program
    {
        static void Main(string[] args)
        {
            string workbookName = "test.xlsx";
            IWorkbook workbook = null;
            ISheet sheet = null;
            IDataValidationHelper dvHelper = null;
            IDataValidationConstraint dvConstraint = null;
            IDataValidation validation = null;
            CellRangeAddressList addressList = null;

            // Using the ss.usermodel allows this class to support both binary
            // and xml based workbooks. The choice of which one to create is
            // made by checking the file extension.
            if (workbookName.EndsWith(".xlsx"))
            {
                workbook = new XSSFWorkbook();
            }
            else
            {
                workbook = new HSSFWorkbook();
            }

            // Build the sheet that will hold the data for the validations. This
            // must be done first as it will create names that are referenced 
            // later.
            sheet = workbook.CreateSheet("Linked Validations");
            BuildDataSheet(sheet);

            // Build the first data validation to occupy cell A1. Note
            // that it retrieves it's data from the named area or region called
            // CHOICES. Further information about this can be found in the
            // static buildDataSheet() method below.
            addressList = new CellRangeAddressList(0, 0, 0, 0);
            dvHelper = sheet.GetDataValidationHelper();
            dvConstraint = dvHelper.CreateFormulaListConstraint("CHOICES");
            validation = dvHelper.CreateValidation(dvConstraint, addressList);
            sheet.AddValidationData(validation);

            // Now, build the linked or dependent drop down list that will
            // occupy cell B1. The key to the whole process is the use of the
            // INDIRECT() function. In the buildDataSheet(0 method, a series of
            // named regions are created and the names of three of them mirror
            // the options available to the user in the first drop down list
            // (in cell A1). Using the INDIRECT() function makes it possible
            // to convert the selection the user makes in that first drop down
            // into the addresses of a named region of cells and then to use
            // those cells to populate the second drop down list.
            addressList = new CellRangeAddressList(0, 0, 1, 1);
            dvConstraint = dvHelper.CreateFormulaListConstraint(
                    "INDIRECT(UPPER($A$1))");
            validation = dvHelper.CreateValidation(dvConstraint, addressList);
            sheet.AddValidationData(validation);

            FileStream sw = File.OpenWrite(workbookName);
            workbook.Write(sw);
            sw.Close();
        }
        /**
     * Called to populate the named areas/regions. The contents of the cells on
     * row one will be used to populate the first drop down list. The contents of
     * the cells on rows two, three and four will be used to populate the second
     * drop down list, just which row will be determined by the choice the user
     * makes in the first drop down list.
     * 
     * In all cases, the approach is to create a row, create and populate cells
     * with data and then specify a name that identifies those cells. With the
     * exception of the first range, the names that are chosen for each range
     * of cells are quite important. In short, each of the options the user 
     * could select in the first drop down list is used as the name for another
     * range of cells. Thus, in this example, the user can select either 
     * 'Animal', 'Vegetable' or 'Mineral' in the first drop down and so the
     * sheet contains ranges named 'ANIMAL', 'VEGETABLE' and 'MINERAL'.
     * 
     * @param dataSheet An instance of a class that implements the Sheet Sheet
     *        interface (HSSFSheet or XSSFSheet).
     */
        static void BuildDataSheet(ISheet dataSheet)
        {
            IRow row = null;
            ICell cell = null;
            IName name = null;

            // The first row will hold the data for the first validation.
            row = dataSheet.CreateRow(10);
            cell = row.CreateCell(0);
            cell.SetCellValue("Animal");
            cell = row.CreateCell(1);
            cell.SetCellValue("Vegetable");
            cell = row.CreateCell(2);
            cell.SetCellValue("Mineral");
            name = dataSheet.Workbook.CreateName();
            name.RefersToFormula = "'Linked Validations'!$A$11:$C$11";
            name.NameName = "CHOICES";

            // The next three rows will hold the data that will be used to
            // populate the second, or linked, drop down list.
            row = dataSheet.CreateRow(11);
            cell = row.CreateCell(0);
            cell.SetCellValue("Lion");
            cell = row.CreateCell(1);
            cell.SetCellValue("Tiger");
            cell = row.CreateCell(2);
            cell.SetCellValue("Leopard");
            cell = row.CreateCell(3);
            cell.SetCellValue("Elephant");
            cell = row.CreateCell(4);
            cell.SetCellValue("Eagle");
            cell = row.CreateCell(5);
            cell.SetCellValue("Horse");
            cell = row.CreateCell(6);
            cell.SetCellValue("Zebra");
            name = dataSheet.Workbook.CreateName();
            name.RefersToFormula = "'Linked Validations'!$A$12:$G$12";
            name.NameName = "ANIMAL";

            row = dataSheet.CreateRow(12);
            cell = row.CreateCell(0);
            cell.SetCellValue("Cabbage");
            cell = row.CreateCell(1);
            cell.SetCellValue("Cauliflower");
            cell = row.CreateCell(2);
            cell.SetCellValue("Potato");
            cell = row.CreateCell(3);
            cell.SetCellValue("Onion");
            cell = row.CreateCell(4);
            cell.SetCellValue("Beetroot");
            cell = row.CreateCell(5);
            cell.SetCellValue("Asparagus");
            cell = row.CreateCell(6);
            cell.SetCellValue("Spinach");
            cell = row.CreateCell(7);
            cell.SetCellValue("Chard");
            name = dataSheet.Workbook.CreateName();
            name.RefersToFormula = "'Linked Validations'!$A$13:$H$13";
            name.NameName = "VEGETABLE";

            row = dataSheet.CreateRow(13);
            cell = row.CreateCell(0);
            cell.SetCellValue("Bauxite");
            cell = row.CreateCell(1);
            cell.SetCellValue("Quartz");
            cell = row.CreateCell(2);
            cell.SetCellValue("Feldspar");
            cell = row.CreateCell(3);
            cell.SetCellValue("Shist");
            cell = row.CreateCell(4);
            cell.SetCellValue("Shale");
            cell = row.CreateCell(5);
            cell.SetCellValue("Mica");
            name = dataSheet.Workbook.CreateName();
            name.RefersToFormula = "'Linked Validations'!$A$14:$F$14";
            name.NameName = "MINERAL";
        }
    }
}
