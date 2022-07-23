using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCases.SS.Formula.Atp
{
    [TestFixture]
    public class TestXMatchFunction
    {
        [Test]
        public void TestMicrosoftExample0()
        {
            HSSFWorkbook wb = initNumWorkbook("Grape");
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            ICell cell = wb.GetSheetAt(0).GetRow(2).CreateCell(5);
            Util.Utils.AssertDouble(fe, cell, "XMATCH(E3,C3:C7)", 2);
            Util.Utils.AssertError(fe, cell, "XMATCH(\"Gra\",C3:C7)", FormulaError.NA);
        }

        [Test]
        public void TestMicrosoftExample1()
        {
            HSSFWorkbook wb = initNumWorkbook("Gra");
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            ICell cell = wb.GetSheetAt(0).GetRow(2).CreateCell(5);
            Util.Utils.AssertDouble(fe, cell, "XMATCH(E3,C3:C7,1)", 2);
            Util.Utils.AssertDouble(fe, cell, "XMATCH(E3,C3:C7,-1)", 5);
            Util.Utils.AssertDouble(fe, cell, "XMATCH(\"Gra\",C3:C7,1)", 2);
            Util.Utils.AssertDouble(fe, cell, "XMATCH(\"Graz\",C3:C7,1)", 3);
            Util.Utils.AssertDouble(fe, cell, "XMATCH(\"Graz\",C3:C7,-1)", 2);
        }

        [Test]
        public void TestMicrosoftExample2()
        {
            //the result in this example is correct but the description seems wrong from my testing
            //the result is based on the position and not a count
            HSSFWorkbook wb = initWorkbook2();
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            ICell cell = wb.GetSheetAt(0).GetRow(3).CreateCell(5);
            Util.Utils.AssertDouble(fe, cell, "XMATCH(F2,C3:C9,1)", 4);
            Util.Utils.AssertDouble(fe, cell, "XMATCH(F2,C3:C9,-1)", 5);
            Util.Utils.AssertError(fe, cell, "XMATCH(F2,C3:C9,2)", FormulaError.NA);
            //Util.Utils.AssertDouble(fe, cell, "XMATCH(35000,C3:C9,1)", 2);
            //Util.Utils.AssertDouble(fe, cell, "XMATCH(36000,C3:C9,1)", 1);
        }

        [Test]
        public void TestMicrosoftExample3()
        {
            HSSFWorkbook wb = initWorkbook3();
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            ICell cell = wb.GetSheetAt(0).GetRow(2).CreateCell(3);
            Util.Utils.AssertDouble(fe, cell, "INDEX(C6:E12,XMATCH(B3,B6:B12),XMATCH(C3,C5:E5))", 8492);

        }

        [Test]
        public void TestMicrosoftExample4()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            ICell cell = wb.CreateSheet().CreateRow(0).CreateCell(0);
            Util.Utils.AssertDouble(fe, cell, "XMATCH(4,{5,4,3,2,1})", 2);
            Util.Utils.AssertDouble(fe, cell, "XMATCH(4.5,{5,4,3,2,1},1)", 1);
        }
        private HSSFWorkbook initNumWorkbook(String lookup)
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            SS.Util.Utils.AddRow(sheet, 0);
            SS.Util.Utils.AddRow(sheet, 1, null, null, "Product", null, "Product", "Position");
            SS.Util.Utils.AddRow(sheet, 2, null, null, "Apple", null, lookup);
            SS.Util.Utils.AddRow(sheet, 3, null, null, "Grape");
            SS.Util.Utils.AddRow(sheet, 4, null, null, "Pear");
            SS.Util.Utils.AddRow(sheet, 5, null, null, "Banana");
            SS.Util.Utils.AddRow(sheet, 6, null, null, "Cherry");
            return wb;
        }

        private HSSFWorkbook initWorkbook2()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            SS.Util.Utils.AddRow(sheet, 0);
            SS.Util.Utils.AddRow(sheet, 1, null, "Sales Rep", "Total Sales", null, "Bonus", 15000);
            SS.Util.Utils.AddRow(sheet, 2, null, "Michael Neipper", 42000);
            SS.Util.Utils.AddRow(sheet, 3, null, "Jan Kotas", 35000);
            SS.Util.Utils.AddRow(sheet, 4, null, "Nancy Freehafer", 25000);
            SS.Util.Utils.AddRow(sheet, 5, null, "Andrew Cencini", 15901);
            SS.Util.Utils.AddRow(sheet, 6, null, "Anne Hellung-Larsen", 13801);
            SS.Util.Utils.AddRow(sheet, 7, null, "Nancy Freehafer", 12181);
            SS.Util.Utils.AddRow(sheet, 8, null, "Mariya Sergienko", 9201);
            return wb;
        }

        private HSSFWorkbook initWorkbook3()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            SS.Util.Utils.AddRow(sheet, 0);
            SS.Util.Utils.AddRow(sheet, 1, null, "Sales Rep", "Month", "Total");
            SS.Util.Utils.AddRow(sheet, 2, null, "Andrew Cencini", "Feb");
            SS.Util.Utils.AddRow(sheet, 3);
            SS.Util.Utils.AddRow(sheet, 4, null, "Sales Rep", "Jan", "Feb", "Mar");
            SS.Util.Utils.AddRow(sheet, 5, null, "Michael Neipper", 3174, 6804, 4713);
            SS.Util.Utils.AddRow(sheet, 6, null, "Jan Kotas", 1656, 8643, 3445);
            SS.Util.Utils.AddRow(sheet, 7, null, "Nancy Freehafer", 2706, 2310, 6606);
            SS.Util.Utils.AddRow(sheet, 8, null, "Andrew Cencini", 4930, 8492, 4474);
            SS.Util.Utils.AddRow(sheet, 9, null, "Anne Hellung-Larsen", 6394, 9846, 4368);
            SS.Util.Utils.AddRow(sheet, 10, null, "Nancy Freehafer", 2539, 8996, 4084);
            SS.Util.Utils.AddRow(sheet, 11, null, "Mariya Sergienko", 4468, 5206, 7343);
            return wb;
        }
    }
}