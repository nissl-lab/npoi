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
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using NUnit.Framework;

namespace NPOI.OOXML.Testcases.XSSF.Streaming
{
    [TestFixture]
    class SXSSFCellTests
    {
        private SXSSFCell _objectToTest;

        [Test]
        public void IfSettingCellErrorValueShouldSetValue()
        {
            _objectToTest = new SXSSFCell(null, CellType.Error);
            _objectToTest.SetCellErrorValue((byte)(0x00));
            Assert.AreEqual((byte)0x00, _objectToTest.ErrorCellValue);
        }

        
        [Test]
        public void IfSettingCellErrorValueAndIsFormulaErrorShouldSetFormulaErrorValue()
        {
            _objectToTest = new SXSSFCell(null, CellType.Numeric);
            _objectToTest.SetCellErrorValue(FormulaError.DIV0.Code);
            Assert.AreEqual((byte)7, _objectToTest.ErrorCellValue);
        }

        [Test]
        public void IfSettingFormulaValueShouldSetFormulaValue()
        {
            _objectToTest = new SXSSFCell(null, CellType.Formula);
            _objectToTest.SetCellFormula("SUM(C4:E4)");
            Assert.AreEqual("SUM(C4:E4)", _objectToTest.CellFormula);
        }

        [Test]
        public void IfGettingCachedFormulaTypeShouldNotThrowErrorIfCellTypeIsNumeric()
        {
            _objectToTest = new SXSSFCell(null, CellType.Formula);
            _objectToTest.SetCellFormula("SUM(C4:E4)");

            Assert.AreEqual(CellType.Numeric, _objectToTest.CachedFormulaResultType);
        }
        [Test]
        public void IfSettingFormulaValueWithNullShouldChangeToBlankCell()
        {
            _objectToTest = new SXSSFCell(null, CellType.Formula);
            _objectToTest.SetCellFormula(null);
            Assert.AreEqual(CellType.Blank, _objectToTest.CellType);
        }

        [Test]
        public void IfCellTypeIsBlankBooleanCellValueShouldReturnFalse()
        {
            _objectToTest = new SXSSFCell(null, CellType.Blank);
            Assert.IsFalse(_objectToTest.BooleanCellValue);
        }
        [Test]
        public void IfCellTypeIsFormulaBooleanCellValueShouldReturnTrueIfValidBooleanFormulaValue()
        {
            _objectToTest = new SXSSFCell(null, CellType.Formula);

            Assert.IsFalse(_objectToTest.BooleanCellValue);
        }

        [Test]
        public void IfGettingCellFormulaShouldReturnFormulaValueIfValidCellType()
        {
            _objectToTest = new SXSSFCell(null, CellType.Formula);
            _objectToTest.SetCellFormula("SUM(C4:E4)");
            Assert.AreEqual("SUM(C4:E4)", _objectToTest.CellFormula);
        }

        [Test]
        public void IfSettingCellFormulaShouldReturnSetForCellTypeFormula()
        {
            _objectToTest = new SXSSFCell(null, CellType.Formula);
            _objectToTest.CellFormula ="SUM(C4:E4)";
            Assert.AreEqual("SUM(C4:E4)", _objectToTest.CellFormula);
        }

        [Test]
        public void IfSettingNullCellFormulaShouldReturnSetCellTypeToBlank()
        {
            _objectToTest = new SXSSFCell(null, CellType.Formula);
            _objectToTest.CellFormula = null;
            Assert.AreEqual(CellType.Blank, _objectToTest.CellType);
        }

        [Test]
        public void IfCellTypeIsBlankNumericCellValueShouldReturnZero()
        {
            _objectToTest = new SXSSFCell(null, CellType.Blank);
            Assert.AreEqual(0.0, _objectToTest.NumericCellValue);
        }

        [Test]
        public void IfCellTypeIsFormulaNumericCellValueShouldReturnPreEvaluatedFormulaValue()
        {
            _objectToTest = new SXSSFCell(null, CellType.Formula);
            _objectToTest.SetCellFormula("SUM(C4:E4)");
            Assert.AreEqual(0.0, _objectToTest.NumericCellValue);
        }

        [Test]
        public void IfCellTypeIsNumericCellValueShouldReturnValue()
        {
            _objectToTest = new SXSSFCell(null, CellType.Formula);
            _objectToTest.SetCellValue(9);
            Assert.AreEqual(9, _objectToTest.NumericCellValue);
        }

    }
}
