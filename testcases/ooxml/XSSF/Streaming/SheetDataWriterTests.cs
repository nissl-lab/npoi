using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using NSubstitute;
using NUnit.Framework;

namespace NPOI.OOXML.Testcases.XSSF.Streaming
{
    [TestFixture]
    class SheetDataWriterTests
    {
        private SheetDataWriter _objectToTest;
        private IRow _row;

        [SetUp]
        public void Init()
        {
            _row = Substitute.For<IRow>();
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
            Assert.AreEqual("<row r=\"" + 1 + "\" customHeight=\"true\"  ht=\"" + row.HeightInPoints + "\">", lines[0]);
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


    }
}
