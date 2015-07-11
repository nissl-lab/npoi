
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

namespace TestCases.SS.Formula.PTG
{

    using System;
    using NUnit.Framework;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.PTG;

    /**
     * Tests for {@link AreaPtg}.
     *
     * @author Dmitriy Kumshayev
     */
    [TestFixture]
    public class TestAreaPtg
    {

        AreaPtg relative;
        AreaPtg absolute;
        [SetUp]
        public void SetUp()
        {
            short firstRow = 5;
            short lastRow = 13;
            short firstCol = 7;
            short lastCol = 17;
            relative = new AreaPtg(firstRow, lastRow, firstCol, lastCol, true, true, true, true);
            absolute = new AreaPtg(firstRow, lastRow, firstCol, lastCol, false, false, false, false);
        }
        [Test]
        public void TestSetColumnsAbsolute()
        {
            resetColumns(absolute);
            validateReference(true, absolute);
        }
        [Test]
        public void TestSetColumnsRelative()
        {
            resetColumns(relative);
            validateReference(false, relative);
        }

        private void validateReference(bool abs, AreaPtg ref1)
        {
            Assert.AreEqual(abs, !ref1.IsFirstColRelative, "First column reference is not " + (abs ? "absolute" : "relative"));
            Assert.AreEqual(abs, !ref1.IsLastColRelative, "Last column reference is not " + (abs ? "absolute" : "relative"));
            Assert.AreEqual(abs, !ref1.IsFirstRowRelative, "First row reference is not " + (abs ? "absolute" : "relative"));
            Assert.AreEqual(abs, !ref1.IsLastRowRelative, "Last row reference is not " + (abs ? "absolute" : "relative"));
        }


        private static void resetColumns(AreaPtg aptg)
        {
            int fc = aptg.FirstColumn;
            int lc = aptg.LastColumn;
            aptg.FirstColumn = (fc);
            aptg.LastColumn = (lc);
            Assert.AreEqual(fc, aptg.FirstColumn);
            Assert.AreEqual(lc, aptg.LastColumn);
        }

        [Test]
        public void TestAbsoluteRelativeRefs()
        {
            AreaPtg sca1 = new AreaPtg(4, 5, 6, 7, true, false, true, false);
            AreaPtg sca2 = new AreaPtg(4, 5, 6, 7, false, true, false, true);
            AreaPtg sca3 = new AreaPtg(5, 5, 7, 7, true, false, true, false);
            AreaPtg sca4 = new AreaPtg(5, 5, 7, 7, false, true, false, true);

            Assert.AreEqual("G5:$H$6", sca1.ToFormulaString(), "first rel., last abs.");
            Assert.AreEqual("$G$5:H6", sca2.ToFormulaString(), "first abs., last rel.");
            Assert.AreEqual("H6:$H$6", sca3.ToFormulaString(), "first rel., last abs.");
            Assert.AreEqual("$H$6:H6", sca4.ToFormulaString(), "first abs., last rel.");

            AreaPtg cla1 = CloneArea(sca1);
            AreaPtg cla2 = CloneArea(sca2);
            AreaPtg cla3 = CloneArea(sca3);
            AreaPtg cla4 = CloneArea(sca4);

            Assert.AreEqual("G5:$H$6", cla1.ToFormulaString(), "first rel., last abs.");
            Assert.AreEqual("$G$5:H6", cla2.ToFormulaString(), "first abs., last rel.");
            Assert.AreEqual("H6:$H$6", cla3.ToFormulaString(), "first rel., last abs.");
            Assert.AreEqual("$H$6:H6", cla4.ToFormulaString(), "first abs., last rel.");
        }
        private AreaPtg CloneArea(AreaPtg a)
        {
            return new AreaPtg(
                    a.FirstRow, a.LastRow, a.FirstColumn, a.LastColumn,
                    a.IsFirstRowRelative, a.IsLastRowRelative, a.IsFirstColRelative, a.IsLastColRelative
            );
        }


        [Test]
        public void TestFormulaParser()
        {
            String formula1 = "SUM($E$5:$E$6)";
            String expectedFormula1 = "SUM($F$5:$F$6)";
            String newFormula1 = ShiftAllColumnsBy1(formula1);
            Assert.AreEqual(expectedFormula1, newFormula1, "Absolute references Changed");

            String formula2 = "SUM(E5:E6)";
            String expectedFormula2 = "SUM(F5:F6)";
            String newFormula2 = ShiftAllColumnsBy1(formula2);
            Assert.AreEqual(expectedFormula2, newFormula2, "Relative references Changed");
        }

        private static String ShiftAllColumnsBy1(String formula)
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