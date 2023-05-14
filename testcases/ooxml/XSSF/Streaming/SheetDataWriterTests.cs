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
using NPOI.XSSF.UserModel;
using NSubstitute;
using NUnit.Framework;
using System.IO;

namespace TestCases.XSSF.Streaming
{
    [TestFixture]
    public class SheetDataWriterTests
    {
        private SheetDataWriter _objectToTest;
        private SXSSFRow _row;
        private ICell _cell;

        [SetUp]
        public void Init()
        {
            var _xssfsheet = Substitute.For<XSSFSheet>();
            var _workbook = Substitute.For<SXSSFWorkbook>();
            var _sheet = Substitute.For<SXSSFSheet>(_workbook, _xssfsheet);
            _row = Substitute.For<SXSSFRow>(_sheet);
            _cell = Substitute.For<ICell>();
        }

        [TearDown]
        public void CleanUp()
        {
            if (_objectToTest != null)
            {
                if (File.Exists(_objectToTest.TemporaryFilePath()))
                {
                    _objectToTest.Dispose();
                    File.Delete(_objectToTest.TemporaryFilePath());
                }
            }

        }
        [Test]
        public void IfCallingEmptyConstructorShouldCreateNonZippedTempFileNonDecoratedStream()
        {
            _objectToTest = new SheetDataWriter();
            Assert.True(_objectToTest.TemporaryFilePath().Contains("poi-sxssf-sheet"));
            Assert.True(!_objectToTest.TemporaryFilePath().Contains(".gz"));
        }

        [Test]
        public void IfWritingRowWithCustomHeightShouldIncludeCustomHeightXml()
        {
            _objectToTest = new SheetDataWriter();
            var row = new SXSSFRow(null);
            row.Height = 1;

            _objectToTest.WriteRow(0, row);
            _objectToTest.Close();

            var lines = File.ReadAllLines(_objectToTest.TemporaryFilePath());

            Assert.True(lines.Length == 2);
            Assert.AreEqual("<row r=\"" + 1 + "\" customHeight=\"true\" ht=\"" + row.HeightInPoints + "\">", lines[0]);
            Assert.AreEqual("</row>", lines[1]);


        }

        [Test]
        public void IfWritingRowWithZeroHeightShouldIncludeHiddenAttributeXml()
        {
            _objectToTest = new SheetDataWriter();

            _row.HasCustomHeight().Returns(false);
            _row.ZeroHeight.Returns(true);

            _objectToTest.WriteRow(0, _row);
            _objectToTest.Close();

            var lines = File.ReadAllLines(_objectToTest.TemporaryFilePath());

            Assert.True(lines.Length == 2);
            Assert.AreEqual("<row r=\"" + 1 + "\" hidden=\"true\">", lines[0]);
            Assert.AreEqual("</row>", lines[1]);


        }

        [Test]
        public void IfWritingRowThatIsFormattedShouldIncludeRowStyleIndexAndCustomFormatAttributeXml()
        {
            _objectToTest = new SheetDataWriter();

            _row.HasCustomHeight().Returns(false);
            _row.ZeroHeight.Returns(false);
            _row.IsFormatted.Returns(true);
            _row.RowStyle.Index.Returns((short)1);
            _objectToTest.WriteRow(0, _row);
            _objectToTest.Close();

            var lines = File.ReadAllLines(_objectToTest.TemporaryFilePath());

            Assert.True(lines.Length == 2);
            Assert.AreEqual("<row r=\"" + 1 + "\" s=\"" + _row.RowStyle.Index + "\" customFormat=\"1\">", lines[0]);
            Assert.AreEqual("</row>", lines[1]);


        }

        [Test]
        public void IfWritingRowHasOutlineLevelGreaterThanZeroShouldAppendOutlineXml()
        {
            _objectToTest = new SheetDataWriter();

            _row.HasCustomHeight().Returns(false);
            _row.ZeroHeight.Returns(false);
            _row.IsFormatted.Returns(false);
            _row.OutlineLevel.Returns(1);
            _objectToTest.WriteRow(0, _row);
            _objectToTest.Close();

            var lines = File.ReadAllLines(_objectToTest.TemporaryFilePath());

            Assert.True(lines.Length == 2);
            Assert.AreEqual("<row r=\"" + 1 + "\" outlineLevel=\"" + _row.OutlineLevel + "\">", lines[0]);
            Assert.AreEqual("</row>", lines[1]);


        }

        [Test]
        public void IfWritingRowIsHiddenShouldAppendHiddenXml()
        {
            _objectToTest = new SheetDataWriter();

            _row.HasCustomHeight().Returns(false);
            _row.ZeroHeight.Returns(false);
            _row.IsFormatted.Returns(false);
            _row.OutlineLevel.Returns(0);
            _row.Hidden.Returns(true);
            _objectToTest.WriteRow(0, _row);
            _objectToTest.Close();

            var lines = File.ReadAllLines(_objectToTest.TemporaryFilePath());

            Assert.True(lines.Length == 2);
            Assert.AreEqual("<row r=\"" + 1 + "\" hidden=\"" + (_row.Hidden.Value ? "1" : "0") + "\">", lines[0]);
            Assert.AreEqual("</row>", lines[1]);


        }

        [Test]
        public void IfWritingRowIsCollapsedShouldAppendCollapsedXml()
        {
            _objectToTest = new SheetDataWriter();

            _row.HasCustomHeight().Returns(false);
            _row.ZeroHeight.Returns(false);
            _row.IsFormatted.Returns(false);
            _row.OutlineLevel.Returns(0);
            _row.Hidden.ReturnsForAnyArgs(x => null);
            _row.Collapsed.Returns(true);
            _objectToTest.WriteRow(0, _row);
            _objectToTest.Close();

            var lines = File.ReadAllLines(_objectToTest.TemporaryFilePath());

            Assert.True(lines.Length == 2);
            Assert.AreEqual("<row r=\"" + 1 + "\" collapsed=\"" + (_row.Collapsed.Value ? "1" : "0") + "\">", lines[0]);
            Assert.AreEqual("</row>", lines[1]);


        }

        [Test]
        public void IfWritingRowWithNoAttributesShouldWriteBasicXml()
        {
            _objectToTest = new SheetDataWriter();

            _row.HasCustomHeight().Returns(false);
            _row.ZeroHeight.Returns(false);
            _row.IsFormatted.Returns(false);
            _row.OutlineLevel.Returns(0);
            _row.Hidden.ReturnsForAnyArgs(x => null);
            _row.Collapsed.Returns(true);
            _objectToTest.WriteRow(0, _row);
            _objectToTest.Close();

            var lines = File.ReadAllLines(_objectToTest.TemporaryFilePath());

            Assert.True(lines.Length == 2);
            Assert.AreEqual("<row r=\"" + 1 + "\" collapsed=\"" + (_row.Collapsed.Value ? "1" : "0") + "\">", lines[0]);
            Assert.AreEqual("</row>", lines[1]);


        }

        [Test]
        public void IfCellTypeIsBlankShouldWriteBlankCellXml()
        {
            _objectToTest = new SheetDataWriter();
            _cell.CellStyle.Index.Returns((short)0);
            _cell.CellType.Returns(CellType.Blank);

            _objectToTest.WriteCell(0, _cell);
            _objectToTest.Close();

            var lines = File.ReadAllLines(_objectToTest.TemporaryFilePath());

            Assert.True(lines.Length == 1);
            Assert.AreEqual("<c r=\"A1\"></c>", lines[0]);

        }

        [Test]
        public void IfCellTypeIsFormulaShouldWriteFormulaCellXml()
        {
            _objectToTest = new SheetDataWriter();
            _cell.CellStyle.Index.Returns((short)0);
            _cell.CellType.Returns(CellType.Formula);
            _cell.CellFormula.Returns("SUM(A1:A3)");
            _cell.GetCachedFormulaResultTypeEnum().Returns(CellType.Numeric);
            _cell.NumericCellValue.Returns(1);

            _objectToTest.WriteCell(0, _cell);
            _objectToTest.Close();

            var lines = File.ReadAllLines(_objectToTest.TemporaryFilePath());

            Assert.True(lines.Length == 1);
            Assert.AreEqual("<c r=\"A1\"><f>SUM(A1:A3)</f><v>1</v></c>", lines[0]);

        }

        [Test]
        public void IfCellTypeIsNumericShouldWriteNumericCellXml()
        {
            _objectToTest = new SheetDataWriter();
            _cell.CellStyle.Index.Returns((short)0);
            _cell.CellType.Returns(CellType.Numeric);
            _cell.NumericCellValue.Returns(1);

            _objectToTest.WriteCell(0, _cell);
            _objectToTest.Close();

            var lines = File.ReadAllLines(_objectToTest.TemporaryFilePath());

            Assert.True(lines.Length == 1);
            Assert.AreEqual("<c r=\"A1\" t=\"n\"><v>1</v></c>", lines[0]);

        }

        [Test]
        public void IfCellTypeIsBooleanTrueShouldWriteBooleanCellTrueXml()
        {
            _objectToTest = new SheetDataWriter();
            _cell.CellStyle.Index.Returns((short)0);
            _cell.CellType.Returns(CellType.Boolean);
            _cell.BooleanCellValue.Returns(true);

            _objectToTest.WriteCell(0, _cell);
            _objectToTest.Close();

            var lines = File.ReadAllLines(_objectToTest.TemporaryFilePath());

            Assert.True(lines.Length == 1);
            Assert.AreEqual("<c r=\"A1\" t=\"b\"><v>1</v></c>", lines[0]);

        }

        [Test]
        public void IfCellTypeIsBooleanFalseShouldWriteBooleanCellFalseXml()
        {
            _objectToTest = new SheetDataWriter();
            _cell.CellStyle.Index.Returns((short)0);
            _cell.CellType.Returns(CellType.Boolean);
            _cell.BooleanCellValue.Returns(false);

            _objectToTest.WriteCell(0, _cell);
            _objectToTest.Close();

            var lines = File.ReadAllLines(_objectToTest.TemporaryFilePath());

            Assert.True(lines.Length == 1);
            Assert.AreEqual("<c r=\"A1\" t=\"b\"><v>0</v></c>", lines[0]);

        }

        [Test]
        public void IfCellTypeIsErrorShouldWriteErrorCellXml()
        {
            _objectToTest = new SheetDataWriter();
            _cell.CellStyle.Index.Returns((short)0);
            _cell.CellType.Returns(CellType.Error);
            _cell.ErrorCellValue.Returns((byte)0x00);

            _objectToTest.WriteCell(0, _cell);
            _objectToTest.Close();

            var lines = File.ReadAllLines(_objectToTest.TemporaryFilePath());

            Assert.True(lines.Length == 1);
            Assert.AreEqual("<c r=\"A1\" t=\"e\"><v>#NULL!</v></c>", lines[0]);

        }

        [Test]
        public void IfCellTypeIsStringShouldWriteStringCellXml()
        {
            _objectToTest = new SheetDataWriter();
            _cell.CellStyle.Index.Returns((short)0);
            _cell.CellType.Returns(CellType.String);
            _cell.StringCellValue.Returns("''<>\t\n\r&\"?         test:SLDFKj    ");

            _objectToTest.WriteCell(0, _cell);
            _objectToTest.Close();

            var lines = File.ReadAllLines(_objectToTest.TemporaryFilePath());

            Assert.True(lines.Length == 1);
            Assert.AreEqual("<c r=\"A1\" t=\"inlineStr\"><is><t xml:space=\"preserve\">\'\'&lt;&gt;&#x9;&#xa;&#xa;&amp;&quot;?         test:SLDFKj    </t></is></c>", lines[0]);

        }

    }
}
