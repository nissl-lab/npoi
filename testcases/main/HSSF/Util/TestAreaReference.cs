/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF.Util
{
    using System;
    using System.IO;
    using NPOI.HSSF.Util;
    //using NPOI.HSSF.Model;
    using NPOI.SS;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.UserModel;

    using NUnit.Framework;using NUnit.Framework.Legacy;

    using TestCases.HSSF;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.HSSF.Model;
    using NPOI.SS.Formula.PTG;

    [TestFixture]
    [Obsolete]
    public class TestAreaReference
    {

        [Test]
        public void TestAreaRef1()
        {
            AreaReference ar = new AreaReference("$A$1:$B$2", SpreadsheetVersion.EXCEL97);
            ClassicAssert.IsFalse(ar.IsSingleCell, "Two cells expected");
            CellReference cf = ar.FirstCell;
            ClassicAssert.IsTrue(cf.Row == 0, "row is1 4");
            ClassicAssert.IsTrue(cf.Col == 0, "col is1 1");
            ClassicAssert.IsTrue(cf.IsRowAbsolute, "row is1 abs");
            ClassicAssert.IsTrue(cf.IsColAbsolute, "col is1 abs");
            ClassicAssert.IsTrue(cf.FormatAsString().Equals("$A$1"), "string is1 $A$1");

            cf = ar.LastCell;
            ClassicAssert.IsTrue(cf.Row == 1, "row is1 4");
            ClassicAssert.IsTrue(cf.Col == 1, "col is1 1");
            ClassicAssert.IsTrue(cf.IsRowAbsolute, "row is1 abs");
            ClassicAssert.IsTrue(cf.IsColAbsolute, "col is1 abs");
            ClassicAssert.IsTrue(cf.FormatAsString().Equals("$B$2"), "string is1 $B$2");

            CellReference[] refs = ar.GetAllReferencedCells();
            ClassicAssert.AreEqual(4, refs.Length);

            ClassicAssert.AreEqual(0, refs[0].Row);
            ClassicAssert.AreEqual(0, refs[0].Col);
            ClassicAssert.IsNull(refs[0].SheetName);

            ClassicAssert.AreEqual(0, refs[1].Row);
            ClassicAssert.AreEqual(1, refs[1].Col);
            ClassicAssert.IsNull(refs[1].SheetName);

            ClassicAssert.AreEqual(1, refs[2].Row);
            ClassicAssert.AreEqual(0, refs[2].Col);
            ClassicAssert.IsNull(refs[2].SheetName);

            ClassicAssert.AreEqual(1, refs[3].Row);
            ClassicAssert.AreEqual(1, refs[3].Col);
            ClassicAssert.IsNull(refs[3].SheetName);
        }

        /**
         * References failed when sheet names were being used
         * Reported by Arne.Clauss@gedas.de
         */
        [Test]
        public void TestReferenceWithSheet()
        {
            AreaReference ar;

            ar = new AreaReference("Tabelle1!B5:B5", SpreadsheetVersion.EXCEL97);
            ClassicAssert.IsTrue(ar.IsSingleCell);
            TestCellReference.ConfirmCell(ar.FirstCell, "Tabelle1", 4, 1, false, false, "Tabelle1!B5");

            ClassicAssert.AreEqual(1, ar.GetAllReferencedCells().Length);


            ar = new AreaReference("Tabelle1!$B$5:$B$7", SpreadsheetVersion.EXCEL97);
            ClassicAssert.IsFalse(ar.IsSingleCell);

            TestCellReference.ConfirmCell(ar.FirstCell, "Tabelle1", 4, 1, true, true, "Tabelle1!$B$5");
            TestCellReference.ConfirmCell(ar.LastCell, "Tabelle1", 6, 1, true, true, "Tabelle1!$B$7");

            // And all that make it up
            CellReference[] allCells = ar.GetAllReferencedCells();
            ClassicAssert.AreEqual(3, allCells.Length);
            TestCellReference.ConfirmCell(allCells[0], "Tabelle1", 4, 1, true, true, "Tabelle1!$B$5");
            TestCellReference.ConfirmCell(allCells[1], "Tabelle1", 5, 1, true, true, "Tabelle1!$B$6");
            TestCellReference.ConfirmCell(allCells[2], "Tabelle1", 6, 1, true, true, "Tabelle1!$B$7");
        }
        [Test]
        public void TestContiguousReferences()
        {
            String refSimple = "$C$10:$C$10";
            String ref2D = "$C$10:$D$11";
            String refDCSimple = "$C$10:$C$10,$D$12:$D$12,$E$14:$E$14";
            String refDC2D = "$C$10:$C$11,$D$12:$D$12,$E$14:$E$20";
            String refDC3D = "Tabelle1!$C$10:$C$14,Tabelle1!$D$10:$D$12";

            // Check that we detect as contiguous properly
            ClassicAssert.IsTrue(AreaReference.IsContiguous(refSimple));
            ClassicAssert.IsTrue(AreaReference.IsContiguous(ref2D));
            ClassicAssert.IsFalse(AreaReference.IsContiguous(refDCSimple));
            ClassicAssert.IsFalse(AreaReference.IsContiguous(refDC2D));
            ClassicAssert.IsFalse(AreaReference.IsContiguous(refDC3D));

            // Check we can only create contiguous entries
            new AreaReference(refSimple, SpreadsheetVersion.EXCEL97);
            new AreaReference(ref2D, SpreadsheetVersion.EXCEL97);
            try
            {
                new AreaReference(refDCSimple, SpreadsheetVersion.EXCEL97);
                Assert.Fail("expected ArgumentException");
            }
            catch (ArgumentException) { }
            try
            {
                new AreaReference(refDC2D, SpreadsheetVersion.EXCEL97);
                Assert.Fail("expected ArgumentException");
            }
            catch (ArgumentException) { }
            try
            {
                new AreaReference(refDC3D, SpreadsheetVersion.EXCEL97);
                Assert.Fail("expected ArgumentException");
            }
            catch (ArgumentException) { }

            // Test that we split as expected
            AreaReference[] refs;

            refs = AreaReference.GenerateContiguous(SpreadsheetVersion.EXCEL97, refSimple);
            ClassicAssert.AreEqual(1, refs.Length);
            ClassicAssert.IsTrue(refs[0].IsSingleCell);
            ClassicAssert.AreEqual("$C$10", refs[0].FormatAsString());

            refs = AreaReference.GenerateContiguous(SpreadsheetVersion.EXCEL97, ref2D);
            ClassicAssert.AreEqual(1, refs.Length);
            ClassicAssert.IsFalse(refs[0].IsSingleCell);
            ClassicAssert.AreEqual("$C$10:$D$11", refs[0].FormatAsString());

            refs = AreaReference.GenerateContiguous(SpreadsheetVersion.EXCEL97, refDCSimple);
            ClassicAssert.AreEqual(3, refs.Length);
            ClassicAssert.IsTrue(refs[0].IsSingleCell);
            ClassicAssert.IsTrue(refs[1].IsSingleCell);
            ClassicAssert.IsTrue(refs[2].IsSingleCell);
            ClassicAssert.AreEqual("$C$10", refs[0].FormatAsString());
            ClassicAssert.AreEqual("$D$12", refs[1].FormatAsString());
            ClassicAssert.AreEqual("$E$14", refs[2].FormatAsString());

            refs = AreaReference.GenerateContiguous(SpreadsheetVersion.EXCEL97, refDC2D);
            ClassicAssert.AreEqual(3, refs.Length);
            ClassicAssert.IsFalse(refs[0].IsSingleCell);
            ClassicAssert.IsTrue(refs[1].IsSingleCell);
            ClassicAssert.IsFalse(refs[2].IsSingleCell);
            ClassicAssert.AreEqual("$C$10:$C$11", refs[0].FormatAsString());
            ClassicAssert.AreEqual("$D$12", refs[1].FormatAsString());
            ClassicAssert.AreEqual("$E$14:$E$20", refs[2].FormatAsString());

            refs = AreaReference.GenerateContiguous(SpreadsheetVersion.EXCEL97, refDC3D);
            ClassicAssert.AreEqual(2, refs.Length);
            ClassicAssert.IsFalse(refs[0].IsSingleCell);
            ClassicAssert.IsFalse(refs[0].IsSingleCell);
            ClassicAssert.AreEqual("Tabelle1!$C$10:$C$14", refs[0].FormatAsString());
            ClassicAssert.AreEqual("Tabelle1!$D$10:$D$12", refs[1].FormatAsString());
            ClassicAssert.AreEqual("Tabelle1", refs[0].FirstCell.SheetName);
            ClassicAssert.AreEqual("Tabelle1", refs[0].LastCell.SheetName);
            ClassicAssert.AreEqual("Tabelle1", refs[1].FirstCell.SheetName);
            ClassicAssert.AreEqual("Tabelle1", refs[1].LastCell.SheetName);
        }
        [Test]
        public void TestDiscontinousReference()
        {
            Stream is1 = HSSFTestDataSamples.OpenSampleFileStream("44167.xls");
            HSSFWorkbook wb = new HSSFWorkbook(is1);
            InternalWorkbook workbook = wb.Workbook;
            HSSFEvaluationWorkbook eb = HSSFEvaluationWorkbook.Create(wb);

            ClassicAssert.AreEqual(1, wb.NumberOfNames);
            String sheetName = "Tabelle1";
            String rawRefA = "$C$10:$C$14";
            String rawRefB = "$C$16:$C$18";
            String refA = sheetName + "!" + rawRefA;
            String refB = sheetName + "!" + rawRefB;
            String ref1 = refA + "," + refB;

            // Check the low level record
            NameRecord nr = workbook.GetNameRecord(0);
            ClassicAssert.IsNotNull(nr);
            ClassicAssert.AreEqual("test", nr.NameText);

            Ptg[] def = nr.NameDefinition;
            ClassicAssert.AreEqual(4, def.Length);

            MemFuncPtg ptgA = (MemFuncPtg)def[0];
            Area3DPtg ptgB = (Area3DPtg)def[1];
            Area3DPtg ptgC = (Area3DPtg)def[2];
            UnionPtg ptgD = (UnionPtg)def[3];
            ClassicAssert.AreEqual("", ptgA.ToFormulaString());
            ClassicAssert.AreEqual(refA, ptgB.ToFormulaString(eb));
            ClassicAssert.AreEqual(refB, ptgC.ToFormulaString(eb));
            ClassicAssert.AreEqual(",", ptgD.ToFormulaString());

            ClassicAssert.AreEqual(ref1, NPOI.HSSF.Model.HSSFFormulaParser.ToFormulaString(wb, nr.NameDefinition));

            // Check the high level definition
            int idx = wb.GetNameIndex("test");
            ClassicAssert.AreEqual(0, idx);
            NPOI.SS.UserModel.IName aNamedCell = wb.GetNameAt(idx);

            // Should have 2 references
            ClassicAssert.AreEqual(ref1, aNamedCell.RefersToFormula);

            // Check the parsing of the reference into cells
            ClassicAssert.IsFalse(AreaReference.IsContiguous(aNamedCell.RefersToFormula));
            AreaReference[] arefs = AreaReference.GenerateContiguous(SpreadsheetVersion.EXCEL97, aNamedCell.RefersToFormula);
            ClassicAssert.AreEqual(2, arefs.Length);
            ClassicAssert.AreEqual(refA, arefs[0].FormatAsString());
            ClassicAssert.AreEqual(refB, arefs[1].FormatAsString());

            for (int i = 0; i < arefs.Length; i++)
            {
                AreaReference ar = arefs[i];
                ConfirmResolveCellRef(wb, ar.FirstCell);
                ConfirmResolveCellRef(wb, ar.LastCell);
            }
        }

        private static void ConfirmResolveCellRef(HSSFWorkbook wb, CellReference cref)
        {
            NPOI.SS.UserModel.ISheet s = wb.GetSheet(cref.SheetName);
            IRow r = s.GetRow(cref.Row);
            ICell c = r.GetCell((int)cref.Col);
            ClassicAssert.IsNotNull(c);
        }
        [Test]
        public void TestSpecialSheetNames()
        {
            AreaReference ar;
            ar = new AreaReference("'Sheet A'!A1:A1", SpreadsheetVersion.EXCEL97);
            ConfirmAreaSheetName(ar, "Sheet A", "'Sheet A'!A1");

            ar = new AreaReference("'Hey! Look Here!'!A1:A1", SpreadsheetVersion.EXCEL97);
            ConfirmAreaSheetName(ar, "Hey! Look Here!", "'Hey! Look Here!'!A1");

            ar = new AreaReference("'O''Toole'!A1:B2", SpreadsheetVersion.EXCEL97);
            ConfirmAreaSheetName(ar, "O'Toole", "'O''Toole'!A1:B2");

            ar = new AreaReference("'one:many'!A1:B2", SpreadsheetVersion.EXCEL97);
            ConfirmAreaSheetName(ar, "one:many", "'one:many'!A1:B2");
        }

        private static void ConfirmAreaSheetName(AreaReference ar, String sheetName, String expectedFullText)
        {
            CellReference[] cells = ar.GetAllReferencedCells();
            ClassicAssert.AreEqual(sheetName, cells[0].SheetName);
            ClassicAssert.AreEqual(expectedFullText, ar.FormatAsString());
        }
        [Test]
        public void TestWholeColumnRefs()
        {
            ConfirmWholeColumnRef("A:A", 0, 0, false, false);
            ConfirmWholeColumnRef("$C:D", 2, 3, true, false);
            ConfirmWholeColumnRef("AD:$AE", 29, 30, false, true);

        }

        private static void ConfirmWholeColumnRef(String ref1, int firstCol, int lastCol, bool firstIsAbs, bool lastIsAbs)
        {
            AreaReference ar = new AreaReference(ref1, SpreadsheetVersion.EXCEL97);
            ConfirmCell(ar.FirstCell, 0, firstCol, true, firstIsAbs);
            ConfirmCell(ar.LastCell, 0xFFFF, lastCol, true, lastIsAbs);
        }

        private static void ConfirmCell(CellReference cell, int row, int col, bool isRowAbs,
                bool isColAbs)
        {
            ClassicAssert.AreEqual(row, cell.Row);
            ClassicAssert.AreEqual(col, cell.Col);
            ClassicAssert.AreEqual(isRowAbs, cell.IsRowAbsolute);
            ClassicAssert.AreEqual(isColAbs, cell.IsColAbsolute);
        }
    }
}