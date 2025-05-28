using NPOI.SS;
using NPOI.SS.Util;
using NUnit.Framework;using NUnit.Framework.Legacy;
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
            ClassicAssert.AreEqual(0, oldStyle.FirstCell.Col);
            ClassicAssert.AreEqual(0, oldStyle.FirstCell.Row);
            ClassicAssert.AreEqual(1, oldStyle.LastCell.Col);
            ClassicAssert.AreEqual(SpreadsheetVersion.EXCEL97.LastRowIndex, oldStyle.LastCell.Row);
            ClassicAssert.IsTrue(oldStyle.IsWholeColumnReference());

            AreaReference oldStyleNonWholeColumn = new AreaReference("A1:B23", SpreadsheetVersion.EXCEL97);
            ClassicAssert.IsFalse(oldStyleNonWholeColumn.IsWholeColumnReference());

            AreaReference newStyle = AreaReference.GetWholeColumn(SpreadsheetVersion.EXCEL2007, "A", "B");
            ClassicAssert.AreEqual(0, newStyle.FirstCell.Col);
            ClassicAssert.AreEqual(0, newStyle.FirstCell.Row);
            ClassicAssert.AreEqual(1, newStyle.LastCell.Col);
            ClassicAssert.AreEqual(SpreadsheetVersion.EXCEL2007.LastRowIndex, newStyle.LastCell.Row);
            ClassicAssert.IsTrue(newStyle.IsWholeColumnReference());

            AreaReference newStyleNonWholeColumn = new AreaReference("A1:B23", SpreadsheetVersion.EXCEL2007);
            ClassicAssert.IsFalse(newStyleNonWholeColumn.IsWholeColumnReference());
        }
        [Test]
        public void TestWholeRow()
        {
            AreaReference oldStyle = AreaReference.GetWholeRow(SpreadsheetVersion.EXCEL97, "1", "2");
            ClassicAssert.AreEqual(0, oldStyle.FirstCell.Col);
            ClassicAssert.AreEqual(0, oldStyle.FirstCell.Row);
            ClassicAssert.AreEqual(SpreadsheetVersion.EXCEL97.LastColumnIndex, oldStyle.LastCell.Col);
            ClassicAssert.AreEqual(1, oldStyle.LastCell.Row);

            AreaReference newStyle = AreaReference.GetWholeRow(SpreadsheetVersion.EXCEL2007, "1", "2");
            ClassicAssert.AreEqual(0, newStyle.FirstCell.Col);
            ClassicAssert.AreEqual(0, newStyle.FirstCell.Row);
            ClassicAssert.AreEqual(SpreadsheetVersion.EXCEL2007.LastColumnIndex, newStyle.LastCell.Col);
            ClassicAssert.AreEqual(1, newStyle.LastCell.Row);
        }

        [Test]
        [Obsolete]
        public void TestFallbackToExcel97IfVersionNotSupplied()
        {
            ClassicAssert.IsTrue(new AreaReference("A:B").IsWholeColumnReference());
            ClassicAssert.IsTrue(AreaReference.IsWholeColumnReference(null, new CellReference("A$1"), new CellReference("A$" + SpreadsheetVersion.EXCEL97.MaxRows)));
        }
    }
}
