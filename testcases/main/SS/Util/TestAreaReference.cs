using NPOI.SS;
using NPOI.SS.Util;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestCases.SS.Util
{
    [TestFixture]
    public class TestAreaReference
    {
        [Test]
        public void TestWholeColumn()
        {
            AreaReference oldStyle = AreaReference.GetWholeColumn(SpreadsheetVersion.EXCEL97, "A", "B");
            Assert.AreEqual(0, oldStyle.FirstCell.Col);
            Assert.AreEqual(0, oldStyle.FirstCell.Row);
            Assert.AreEqual(1, oldStyle.LastCell.Col);
            Assert.AreEqual(SpreadsheetVersion.EXCEL97.LastRowIndex, oldStyle.LastCell.Row);
            Assert.IsTrue(oldStyle.IsWholeColumnReference());

            AreaReference oldStyleNonWholeColumn = new AreaReference("A1:B23", SpreadsheetVersion.EXCEL97);
            Assert.IsFalse(oldStyleNonWholeColumn.IsWholeColumnReference());

            AreaReference newStyle = AreaReference.GetWholeColumn(SpreadsheetVersion.EXCEL2007, "A", "B");
            Assert.AreEqual(0, newStyle.FirstCell.Col);
            Assert.AreEqual(0, newStyle.FirstCell.Row);
            Assert.AreEqual(1, newStyle.LastCell.Col);
            Assert.AreEqual(SpreadsheetVersion.EXCEL2007.LastRowIndex, newStyle.LastCell.Row);
            Assert.IsTrue(newStyle.IsWholeColumnReference());

            AreaReference newStyleNonWholeColumn = new AreaReference("A1:B23", SpreadsheetVersion.EXCEL2007);
            Assert.IsFalse(newStyleNonWholeColumn.IsWholeColumnReference());
        }
        [Test]
        public void TestWholeRow()
        {
            AreaReference oldStyle = AreaReference.GetWholeRow(SpreadsheetVersion.EXCEL97, "1", "2");
            Assert.AreEqual(0, oldStyle.FirstCell.Col);
            Assert.AreEqual(0, oldStyle.FirstCell.Row);
            Assert.AreEqual(SpreadsheetVersion.EXCEL97.LastColumnIndex, oldStyle.LastCell.Col);
            Assert.AreEqual(1, oldStyle.LastCell.Row);

            AreaReference newStyle = AreaReference.GetWholeRow(SpreadsheetVersion.EXCEL2007, "1", "2");
            Assert.AreEqual(0, newStyle.FirstCell.Col);
            Assert.AreEqual(0, newStyle.FirstCell.Row);
            Assert.AreEqual(SpreadsheetVersion.EXCEL2007.LastColumnIndex, newStyle.LastCell.Col);
            Assert.AreEqual(1, newStyle.LastCell.Row);
        }

        [Test]
        [Obsolete]
        public void TestFallbackToExcel97IfVersionNotSupplied()
        {
            Assert.IsTrue(new AreaReference("A:B").IsWholeColumnReference());
            Assert.IsTrue(AreaReference.IsWholeColumnReference(null, new CellReference("A$1"), new CellReference("A$" + SpreadsheetVersion.EXCEL97.MaxRows)));
        }
    }
}
