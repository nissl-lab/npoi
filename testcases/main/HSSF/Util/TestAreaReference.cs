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
    using NPOI.SS.Formula;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.UserModel;

    using NUnit.Framework;

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
            AreaReference ar = new AreaReference("$A$1:$B$2");
            Assert.IsFalse(ar.IsSingleCell, "Two cells expected");
            CellReference cf = ar.FirstCell;
            Assert.IsTrue(cf.Row == 0, "row is1 4");
            Assert.IsTrue(cf.Col == 0, "col is1 1");
            Assert.IsTrue(cf.IsRowAbsolute, "row is1 abs");
            Assert.IsTrue(cf.IsColAbsolute, "col is1 abs");
            Assert.IsTrue(cf.FormatAsString().Equals("$A$1"), "string is1 $A$1");

            cf = ar.LastCell;
            Assert.IsTrue(cf.Row == 1, "row is1 4");
            Assert.IsTrue(cf.Col == 1, "col is1 1");
            Assert.IsTrue(cf.IsRowAbsolute, "row is1 abs");
            Assert.IsTrue(cf.IsColAbsolute, "col is1 abs");
            Assert.IsTrue(cf.FormatAsString().Equals("$B$2"), "string is1 $B$2");

            CellReference[] refs = ar.GetAllReferencedCells();
            Assert.AreEqual(4, refs.Length);

            Assert.AreEqual(0, refs[0].Row);
            Assert.AreEqual(0, refs[0].Col);
            Assert.IsNull(refs[0].SheetName);

            Assert.AreEqual(0, refs[1].Row);
            Assert.AreEqual(1, refs[1].Col);
            Assert.IsNull(refs[1].SheetName);

            Assert.AreEqual(1, refs[2].Row);
            Assert.AreEqual(0, refs[2].Col);
            Assert.IsNull(refs[2].SheetName);

            Assert.AreEqual(1, refs[3].Row);
            Assert.AreEqual(1, refs[3].Col);
            Assert.IsNull(refs[3].SheetName);
        }

        /**
         * References failed when sheet names were being used
         * Reported by Arne.Clauss@gedas.de
         */
        [Test]
        public void TestReferenceWithSheet()
        {
            AreaReference ar;

            ar = new AreaReference("Tabelle1!B5:B5");
            Assert.IsTrue(ar.IsSingleCell);
            TestCellReference.ConfirmCell(ar.FirstCell, "Tabelle1", 4, 1, false, false, "Tabelle1!B5");

            Assert.AreEqual(1, ar.GetAllReferencedCells().Length);


            ar = new AreaReference("Tabelle1!$B$5:$B$7");
            Assert.IsFalse(ar.IsSingleCell);

            TestCellReference.ConfirmCell(ar.FirstCell, "Tabelle1", 4, 1, true, true, "Tabelle1!$B$5");
            TestCellReference.ConfirmCell(ar.LastCell, "Tabelle1", 6, 1, true, true, "Tabelle1!$B$7");

            // And all that make it up
            CellReference[] allCells = ar.GetAllReferencedCells();
            Assert.AreEqual(3, allCells.Length);
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
            Assert.IsTrue(AreaReference.IsContiguous(refSimple));
            Assert.IsTrue(AreaReference.IsContiguous(ref2D));
            Assert.IsFalse(AreaReference.IsContiguous(refDCSimple));
            Assert.IsFalse(AreaReference.IsContiguous(refDC2D));
            Assert.IsFalse(AreaReference.IsContiguous(refDC3D));

            // Check we can only create contiguous entries
            new AreaReference(refSimple);
            new AreaReference(ref2D);
            try
            {
                new AreaReference(refDCSimple);
                Assert.Fail();
            }
            catch (ArgumentException) { }
            try
            {
                new AreaReference(refDC2D);
                Assert.Fail();
            }
            catch (ArgumentException) { }
            try
            {
                new AreaReference(refDC3D);
                Assert.Fail();
            }
            catch (ArgumentException) { }

            // Test that we split as expected
            AreaReference[] refs;

            refs = AreaReference.GenerateContiguous(refSimple);
            Assert.AreEqual(1, refs.Length);
            Assert.IsTrue(refs[0].IsSingleCell);
            Assert.AreEqual("$C$10", refs[0].FormatAsString());

            refs = AreaReference.GenerateContiguous(ref2D);
            Assert.AreEqual(1, refs.Length);
            Assert.IsFalse(refs[0].IsSingleCell);
            Assert.AreEqual("$C$10:$D$11", refs[0].FormatAsString());

            refs = AreaReference.GenerateContiguous(refDCSimple);
            Assert.AreEqual(3, refs.Length);
            Assert.IsTrue(refs[0].IsSingleCell);
            Assert.IsTrue(refs[1].IsSingleCell);
            Assert.IsTrue(refs[2].IsSingleCell);
            Assert.AreEqual("$C$10", refs[0].FormatAsString());
            Assert.AreEqual("$D$12", refs[1].FormatAsString());
            Assert.AreEqual("$E$14", refs[2].FormatAsString());

            refs = AreaReference.GenerateContiguous(refDC2D);
            Assert.AreEqual(3, refs.Length);
            Assert.IsFalse(refs[0].IsSingleCell);
            Assert.IsTrue(refs[1].IsSingleCell);
            Assert.IsFalse(refs[2].IsSingleCell);
            Assert.AreEqual("$C$10:$C$11", refs[0].FormatAsString());
            Assert.AreEqual("$D$12", refs[1].FormatAsString());
            Assert.AreEqual("$E$14:$E$20", refs[2].FormatAsString());

            refs = AreaReference.GenerateContiguous(refDC3D);
            Assert.AreEqual(2, refs.Length);
            Assert.IsFalse(refs[0].IsSingleCell);
            Assert.IsFalse(refs[0].IsSingleCell);
            Assert.AreEqual("Tabelle1!$C$10:$C$14", refs[0].FormatAsString());
            Assert.AreEqual("Tabelle1!$D$10:$D$12", refs[1].FormatAsString());
            Assert.AreEqual("Tabelle1", refs[0].FirstCell.SheetName);
            Assert.AreEqual("Tabelle1", refs[0].LastCell.SheetName);
            Assert.AreEqual("Tabelle1", refs[1].FirstCell.SheetName);
            Assert.AreEqual("Tabelle1", refs[1].LastCell.SheetName);
        }
        [Test]
        public void TestDiscontinousReference()
        {
            Stream is1 = HSSFTestDataSamples.OpenSampleFileStream("44167.xls");
            HSSFWorkbook wb = new HSSFWorkbook(is1);
            InternalWorkbook workbook = wb.Workbook;
            HSSFEvaluationWorkbook eb = HSSFEvaluationWorkbook.Create(wb);

            Assert.AreEqual(1, wb.NumberOfNames);
            String sheetName = "Tabelle1";
            String rawRefA = "$C$10:$C$14";
            String rawRefB = "$C$16:$C$18";
            String refA = sheetName + "!" + rawRefA;
            String refB = sheetName + "!" + rawRefB;
            String ref1 = refA + "," + refB;

            // Check the low level record
            NameRecord nr = workbook.GetNameRecord(0);
            Assert.IsNotNull(nr);
            Assert.AreEqual("test", nr.NameText);

            Ptg[] def = nr.NameDefinition;
            Assert.AreEqual(4, def.Length);

            MemFuncPtg ptgA = (MemFuncPtg)def[0];
            Area3DPtg ptgB = (Area3DPtg)def[1];
            Area3DPtg ptgC = (Area3DPtg)def[2];
            UnionPtg ptgD = (UnionPtg)def[3];
            Assert.AreEqual("", ptgA.ToFormulaString());
            Assert.AreEqual(refA, ptgB.ToFormulaString(eb));
            Assert.AreEqual(refB, ptgC.ToFormulaString(eb));
            Assert.AreEqual(",", ptgD.ToFormulaString());

            Assert.AreEqual(ref1, NPOI.HSSF.Model.HSSFFormulaParser.ToFormulaString(wb, nr.NameDefinition));

            // Check the high level definition
            int idx = wb.GetNameIndex("test");
            Assert.AreEqual(0, idx);
            NPOI.SS.UserModel.IName aNamedCell = wb.GetNameAt(idx);

            // Should have 2 references
            Assert.AreEqual(ref1, aNamedCell.RefersToFormula);

            // Check the parsing of the reference into cells
            Assert.IsFalse(AreaReference.IsContiguous(aNamedCell.RefersToFormula));
            AreaReference[] arefs = AreaReference.GenerateContiguous(aNamedCell.RefersToFormula);
            Assert.AreEqual(2, arefs.Length);
            Assert.AreEqual(refA, arefs[0].FormatAsString());
            Assert.AreEqual(refB, arefs[1].FormatAsString());

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
            Assert.IsNotNull(c);
        }
        [Test]
        public void TestSpecialSheetNames()
        {
            AreaReference ar;
            ar = new AreaReference("'Sheet A'!A1:A1");
            ConfirmAreaSheetName(ar, "Sheet A", "'Sheet A'!A1");

            ar = new AreaReference("'Hey! Look Here!'!A1:A1");
            ConfirmAreaSheetName(ar, "Hey! Look Here!", "'Hey! Look Here!'!A1");

            ar = new AreaReference("'O''Toole'!A1:B2");
            ConfirmAreaSheetName(ar, "O'Toole", "'O''Toole'!A1:B2");

            ar = new AreaReference("'one:many'!A1:B2");
            ConfirmAreaSheetName(ar, "one:many", "'one:many'!A1:B2");
        }

        private static void ConfirmAreaSheetName(AreaReference ar, String sheetName, String expectedFullText)
        {
            CellReference[] cells = ar.GetAllReferencedCells();
            Assert.AreEqual(sheetName, cells[0].SheetName);
            Assert.AreEqual(expectedFullText, ar.FormatAsString());
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
            AreaReference ar = new AreaReference(ref1);
            ConfirmCell(ar.FirstCell, 0, firstCol, true, firstIsAbs);
            ConfirmCell(ar.LastCell, 0xFFFF, lastCol, true, lastIsAbs);
        }

        private static void ConfirmCell(CellReference cell, int row, int col, bool isRowAbs,
                bool isColAbs)
        {
            Assert.AreEqual(row, cell.Row);
            Assert.AreEqual(col, cell.Col);
            Assert.AreEqual(isRowAbs, cell.IsRowAbsolute);
            Assert.AreEqual(isColAbs, cell.IsColAbsolute);
        }
    }
}