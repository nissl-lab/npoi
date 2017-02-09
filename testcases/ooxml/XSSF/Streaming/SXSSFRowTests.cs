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
    class SXSSFRowTests
    {
        //TODO: Add Tests for creating cells.
        private SXSSFRow _objectToTest;

        [SetUp]
        public void Init()
        {
            _objectToTest = new SXSSFRow(null);
        }
        [Test]
        public void IfCreatingCellShouldReturnBlankCell()
        {
            var result = _objectToTest.CreateCell(0);
            Assert.AreEqual(CellType.Blank, result.CellType);
        }

        [Test]
        public void IfCreatingCellWithTypeBooleanShouldReturnCellofTypeBoolean()
        {
            var result = _objectToTest.CreateCell(0, CellType.Boolean);
            Assert.AreEqual(CellType.Boolean, result.CellType);
        }

        [Test]
        public void IfCreatingCellWithTypeFormulaShouldReturnCellofTypeFormula()
        {
            var result = _objectToTest.CreateCell(0, CellType.Formula);
            Assert.AreEqual(CellType.Formula, result.CellType);
        }

        [Test]
        public void IfCreatingCellWithTypeErrorShouldReturnCellofTypeError()
        {
            var result = _objectToTest.CreateCell(0, CellType.Error);
            Assert.AreEqual(CellType.Error, result.CellType);
        }

        [Test]
        public void IfCreatingCellWithTypeNumericShouldReturnCellofTypeNumeric()
        {
            var result = _objectToTest.CreateCell(0, CellType.Numeric);
            Assert.AreEqual(CellType.Numeric, result.CellType);
        }

        [Test]
        public void IfCreatingCellWithTypeStringShouldReturnCellofTypeString()
        {
            var result = _objectToTest.CreateCell(0, CellType.String);
            Assert.AreEqual(CellType.String, result.CellType);
        }

        //TODO add test for cell out of bounds.
    }
}
