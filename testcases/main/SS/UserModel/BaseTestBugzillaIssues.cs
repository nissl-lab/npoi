/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.SS.usermodel;

using junit.framework.TestCase;
using NPOI.SS.ITestDataProvider;
using NPOI.SS.SpreadsheetVersion;
using NPOI.SS.util.CellRangeAddress;

/**
 * A base class for bugzilla issues that can be described in terms of common ss interfaces.
 *
 * @author Yegor Kozlov
 */
public abstract class BaseTestBugzillaIssues : TestCase {

    protected abstract ITestDataProvider GetTestDataProvider();

    /**
     *
     * Test writing a hyperlink
     * Open resulting sheet in Excel and check that A1 Contains a hyperlink
     *
     * Also tests bug 15353 (problems with hyperlinks to Google)
     */
    public void test23094() {
        Workbook wb = GetTestDataProvider().CreateWorkbook();
        Sheet s = wb.CreateSheet();
        Row r = s.CreateRow(0);
        r.CreateCell(0).SetCellFormula("HYPERLINK(\"http://jakarta.apache.org\",\"Jakarta\")");
        r.CreateCell(1).SetCellFormula("HYPERLINK(\"http://google.com\",\"Google\")");

        wb = GetTestDataProvider().WriteOutAndReadBack(wb);
        r = wb.GetSheetAt(0).GetRow(0);

        Cell cell_0 = r.GetCell(0);
        Assert.AreEqual("HYPERLINK(\"http://jakarta.apache.org\",\"Jakarta\")", cell_0.GetCellFormula());
        Cell cell_1 = r.GetCell(1);
        Assert.AreEqual("HYPERLINK(\"http://google.com\",\"Google\")", cell_1.GetCellFormula());
    }

    /**
     * test writing a file with large number of unique strings,
     * open resulting file in Excel to check results!
     * @param  num the number of strings to generate
     */
    public void baseTest15375(int num) {
        Workbook wb = GetTestDataProvider().CreateWorkbook();
        Sheet sheet = wb.CreateSheet();
        CreationHelper factory = wb.GetCreationHelper();

        String tmp1 = null;
        String tmp2 = null;
        String tmp3 = null;

        for (int i = 0; i < num; i++) {
            tmp1 = "Test1" + i;
            tmp2 = "Test2" + i;
            tmp3 = "Test3" + i;

            Row row = sheet.CreateRow(i);

            Cell cell = row.CreateCell(0);
            cell.SetCellValue(factory.CreateRichTextString(tmp1));
            cell = row.CreateCell(1);
            cell.SetCellValue(factory.CreateRichTextString(tmp2));
            cell = row.CreateCell(2);
            cell.SetCellValue(factory.CreateRichTextString(tmp3));
        }
        wb = GetTestDataProvider().WriteOutAndReadBack(wb);
        for (int i = 0; i < num; i++) {
            tmp1 = "Test1" + i;
            tmp2 = "Test2" + i;
            tmp3 = "Test3" + i;

            Row row = sheet.GetRow(i);

            Assert.AreEqual(tmp1, row.GetCell(0).GetStringCellValue());
            Assert.AreEqual(tmp2, row.GetCell(1).GetStringCellValue());
            Assert.AreEqual(tmp3, row.GetCell(2).GetStringCellValue());
        }
    }

    /**
     * Merged regions were being Removed from the parent in Cloned sheets
     */
    public void test22720() {
       Workbook workBook = GetTestDataProvider().CreateWorkbook();
       workBook.CreateSheet("TEST");
       Sheet template = workBook.GetSheetAt(0);

       template.AddMergedRegion(new CellRangeAddress(0, 1, 0, 2));
       template.AddMergedRegion(new CellRangeAddress(1, 2, 0, 2));

       Sheet clone = workBook.CloneSheet(0);
       int originalMerged = template.GetNumMergedRegions();
       Assert.AreEqual("2 merged regions", 2, originalMerged);

       //remove merged regions from clone
       for (int i=template.GetNumMergedRegions()-1; i>=0; i--) {
           Clone.RemoveMergedRegion(i);
       }

       Assert.AreEqual("Original Sheet's Merged Regions were Removed", originalMerged, template.GetNumMergedRegions());
       //check if template's merged regions are OK
       if (template.GetNumMergedRegions()>0) {
            // fetch the first merged region...EXCEPTION OCCURS HERE
            template.GetMergedRegion(0);
       }
       //make sure we dont exception

    }

    public void test28031() {
        Workbook wb = GetTestDataProvider().CreateWorkbook();
        Sheet sheet = wb.CreateSheet();
        wb.SetSheetName(0, "Sheet1");

        Row row = sheet.CreateRow(0);
        Cell cell = row.CreateCell(0);
        String formulaText =
            "IF(ROUND(A2*B2*C2,2)>ROUND(B2*D2,2),ROUND(A2*B2*C2,2),ROUND(B2*D2,2))";
        cell.SetCellFormula(formulaText);

        Assert.AreEqual(formulaText, cell.GetCellFormula());
        wb = GetTestDataProvider().WriteOutAndReadBack(wb);
        cell = wb.GetSheetAt(0).GetRow(0).GetCell(0);
        Assert.AreEqual("IF(ROUND(A2*B2*C2,2)>ROUND(B2*D2,2),ROUND(A2*B2*C2,2),ROUND(B2*D2,2))", cell.GetCellFormula());
    }

    /**
     * Bug 21334: "File error: data may have been lost" with a file
     * that Contains macros and this formula:
     * {=SUM(IF(FREQUENCY(IF(LEN(V4:V220)>0,MATCH(V4:V220,V4:V220,0),""),IF(LEN(V4:V220)>0,MATCH(V4:V220,V4:V220,0),""))>0,1))}
     */
    public void test21334() {
        Workbook wb = GetTestDataProvider().CreateWorkbook();
        Sheet sh = wb.CreateSheet();
        Cell cell = sh.CreateRow(0).CreateCell(0);
        String formula = "SUM(IF(FREQUENCY(IF(LEN(V4:V220)>0,MATCH(V4:V220,V4:V220,0),\"\"),IF(LEN(V4:V220)>0,MATCH(V4:V220,V4:V220,0),\"\"))>0,1))";
        cell.SetCellFormula(formula);

        Workbook wb_sv = GetTestDataProvider().WriteOutAndReadBack(wb);
        Cell cell_sv = wb_sv.GetSheetAt(0).GetRow(0).GetCell(0);
        Assert.AreEqual(formula, cell_sv.GetCellFormula());
    }

    /** another test for the number of unique strings issue
     *test opening the resulting file in Excel*/
    public void test22568() {
        int r=2000;int c=3;

        Workbook wb = GetTestDataProvider().CreateWorkbook();
        Sheet sheet = wb.CreateSheet("ExcelTest") ;

        int col_cnt=0, rw_cnt=0 ;

        col_cnt = c;
        rw_cnt = r;

        Row rw ;
        rw = sheet.CreateRow(0) ;
        //Header row
        for(int j=0; j<col_cnt; j++){
            Cell cell = rw.CreateCell(j) ;
            cell.SetCellValue("Col " + (j+1));
        }

        for(int i=1; i<rw_cnt; i++){
            rw = sheet.CreateRow(i) ;
            for(int j=0; j<col_cnt; j++){
                Cell cell = rw.CreateCell(j) ;
                cell.SetCellValue("Row:" + (i+1) + ",Column:" + (j+1));
            }
        }

        sheet.SetDefaultColumnWidth(18) ;

        wb = GetTestDataProvider().WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        rw = sheet.GetRow(0);
        //Header row
        for(int j=0; j<col_cnt; j++){
            Cell cell = rw.GetCell(j) ;
            Assert.AreEqual("Col " + (j+1), cell.GetStringCellValue());
        }
        for(int i=1; i<rw_cnt; i++){
            rw = sheet.GetRow(i) ;
            for(int j=0; j<col_cnt; j++){
                Cell cell = rw.GetCell(j) ;
                Assert.AreEqual("Row:" + (i+1) + ",Column:" + (j+1), cell.GetStringCellValue());
            }
        }
    }

    /**
     * Bug 42448: Can't parse SUMPRODUCT(A!C7:A!C67, B8:B68) / B69
     */
    public void test42448(){
        Workbook wb = GetTestDataProvider().CreateWorkbook();
        Cell cell = wb.CreateSheet().CreateRow(0).CreateCell(0);
        cell.SetCellFormula("SUMPRODUCT(A!C7:A!C67, B8:B68) / B69");
        Assert.IsTrue("no errors parsing formula", true);
    }

    public void test18800() {
       Workbook book = GetTestDataProvider().CreateWorkbook();
       book.CreateSheet("TEST");
       Sheet sheet = book.CloneSheet(0);
       book.SetSheetName(1,"CLONE");
       sheet.CreateRow(0).CreateCell(0).SetCellValue("Test");

       book = GetTestDataProvider().WriteOutAndReadBack(book);
       sheet = book.GetSheet("CLONE");
       Row row = sheet.GetRow(0);
       Cell cell = row.GetCell(0);
       Assert.AreEqual("Test", cell.GetRichStringCellValue().GetString());
   }

    private static void AddNewSheetWithCellsA1toD4(Workbook book, int sheet) {

        Sheet sht = book .CreateSheet("s" + sheet);
        for (int r=0; r < 4; r++) {

            Row   row = sht.CreateRow (r);
            for (int c=0; c < 4; c++) {

                Cell cel = row.CreateCell(c);
                cel.SetCellValue(sheet*100 + r*10 + c);
            }
        }
    }

    public void testBug43093() {
        Workbook xlw = GetTestDataProvider().CreateWorkbook();

        AddNewSheetWithCellsA1toD4(xlw, 1);
        AddNewSheetWithCellsA1toD4(xlw, 2);
        AddNewSheetWithCellsA1toD4(xlw, 3);
        AddNewSheetWithCellsA1toD4(xlw, 4);

        Sheet s2   = xlw.GetSheet("s2");
        Row   s2r3 = s2.GetRow(3);
        Cell  s2E4 = s2r3.CreateCell(4);
        s2E4.SetCellFormula("SUM(s3!B2:C3)");

        FormulaEvaluator eva = xlw.GetCreationHelper().CreateFormulaEvaluator();
        double d = eva.Evaluate(s2E4).GetNumberValue();

        Assert.AreEqual(d, (311+312+321+322), 0.0000001);
    }

    public void testMaxFunctionArguments_bug46729(){
        String[] func = {"COUNT", "AVERAGE", "MAX", "MIN", "OR", "SUBTOTAL", "SKEW"};

        SpreadsheetVersion ssVersion = GetTestDataProvider().GetSpreadsheetVersion();
        Workbook wb = GetTestDataProvider().CreateWorkbook();
        Cell cell = wb.CreateSheet().CreateRow(0).CreateCell(0);

        String fmla;
        for (String name : func) {

            fmla = CreateFunction(name, 5);
            cell.SetCellFormula(fmla);

            fmla = CreateFunction(name, ssVersion.GetMaxFunctionArgs());
            cell.SetCellFormula(fmla);

            try {
                fmla = CreateFunction(name, ssVersion.GetMaxFunctionArgs() + 1);
                cell.SetCellFormula(fmla);
                fail("Expected FormulaParseException");
            } catch (Exception e){
                 Assert.IsTrue(e.GetMessage().startsWith("Too many arguments to function '"+name+"'"));
            }
        }
    }

    private String CreateFunction(String name, int maxArgs){
        StringBuilder fmla = new StringBuilder();
        fmla.Append(name);
        fmla.Append("(");
        for(int i=0; i < maxArgs; i++){
            if(i > 0) fmla.Append(',');
            fmla.Append("A1");
        }
        fmla.Append(")");
        return fmla.ToString();
    }
}





