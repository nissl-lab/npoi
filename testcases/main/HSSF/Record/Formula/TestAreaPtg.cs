        
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

namespace TestCases.HSSF.Record.Formula
{
    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record.Formula;
    using NPOI.HSSF.UserModel;


    /**
     * Tests for {@link AreaPtg}.
     *
     * @author Dmitriy Kumshayev
     */
    [TestClass]
    public class TestAreaPtg
    {

        AreaPtg relative;
        AreaPtg absolute;

        [TestInitialize]
        public void SetUp()
        {
            short firstRow = 5;
            short lastRow = 13;
            short firstCol = 7;
            short lastCol = 17;
            relative = new AreaPtg(firstRow, lastRow, firstCol, lastCol, true, true, true, true);
            absolute = new AreaPtg(firstRow, lastRow, firstCol, lastCol, false, false, false, false);
        }
        [TestMethod]
        public void TestSetColumnsAbsolute()
        {
            ResetColumns(absolute);
            ValidateReference(true, absolute);
        }
        [TestMethod]
        public void TestSetColumnsRelative()
        {
            ResetColumns(relative);
            ValidateReference(false, relative);
        }

        private void ValidateReference(bool abs, AreaPtg ref1)
        {
            Assert.AreEqual(abs, !ref1.IsFirstColRelative, "First column reference is1 not " + (abs ? "absolute" : "relative"));
            Assert.AreEqual(abs, !ref1.IsLastColRelative, "Last column reference is1 not " + (abs ? "absolute" : "relative"));
            Assert.AreEqual(abs, !ref1.IsFirstRowRelative, "First row reference is1 not " + (abs ? "absolute" : "relative"));
            Assert.AreEqual(abs, !ref1.IsLastRowRelative, "Last row reference is1 not " + (abs ? "absolute" : "relative"));
        }

        
        private static void ResetColumns(AreaPtg aptg)
        {
            int fc = aptg.FirstColumn;
            int lc = aptg.LastColumn;
            aptg.FirstColumn = (fc);
            aptg.LastColumn = (lc);
            Assert.AreEqual(fc, aptg.FirstColumn);
            Assert.AreEqual(lc, aptg.LastColumn);
        }
        [TestMethod]
        public void TestFormulaParser()
        {
            String formula1 = "SUM($E$5:$E$6)";
            String expectedFormula1 = "SUM($F$5:$F$6)";
            String newFormula1 = ShiftAllColumnsBy1(formula1);
            Assert.AreEqual(expectedFormula1, newFormula1, "Absolute references changed");

            String formula2 = "SUM(E5:E6)";
            String expectedFormula2 = "SUM(F5:F6)";
            String newFormula2 = ShiftAllColumnsBy1(formula2);
            Assert.AreEqual(expectedFormula2, newFormula2, "Relative references changed");
        }

        private String ShiftAllColumnsBy1(String formula)
        {
            int letUsShiftColumn1By1Column = 1;
            HSSFWorkbook wb = null;

            Ptg[] ptgs = HSSFFormulaParser.Parse(formula, wb);
            for (int i = 0; i < ptgs.Length; i++)
            {
                Ptg ptg = ptgs[i];
                if (ptg is AreaPtg)
                {
                    AreaPtg aptg = (AreaPtg)ptg;
                    aptg.FirstColumn = ((short)(aptg.FirstColumn + letUsShiftColumn1By1Column));
                    aptg.LastColumn = ((short)(aptg.LastColumn + letUsShiftColumn1By1Column));
                }
            }
            String newFormula = HSSFFormulaParser.ToFormulaString(wb, ptgs);
            return newFormula;
        }


    }
}
