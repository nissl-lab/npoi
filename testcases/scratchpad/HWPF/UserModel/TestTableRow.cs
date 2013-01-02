using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HWPF;
using NPOI.HWPF.UserModel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TestCases.HWPF.UserModel
{
    [TestClass]
    public class TestTableRow
    {
        [TestMethod]
        public void TestInnerTableCellsDetection()
        {
            HWPFDocument hwpfDocument = new HWPFDocument(POIDataSamples
                    .GetDocumentInstance().OpenResourceAsStream("innertable.doc"));
            hwpfDocument.GetRange();

            Range documentRange = hwpfDocument.GetRange();
            Paragraph startOfInnerTable = documentRange.GetParagraph(6);

            Table innerTable = documentRange.GetTable(startOfInnerTable);
            Assert.AreEqual(2, innerTable.NumRows);

            TableRow tableRow = innerTable.GetRow(0);
            Assert.AreEqual(2, tableRow.NumCells());
        }
        [TestMethod]
        public void TestOuterTableCellsDetection()
        {
            HWPFDocument hwpfDocument = new HWPFDocument(POIDataSamples
                    .GetDocumentInstance().OpenResourceAsStream("innertable.doc"));
            hwpfDocument.GetRange();

            Range documentRange = hwpfDocument.GetRange();
            Paragraph startOfOuterTable = documentRange.GetParagraph(0);

            Table outerTable = documentRange.GetTable(startOfOuterTable);
            Assert.AreEqual(3, outerTable.NumRows);

            Assert.AreEqual(3, outerTable.GetRow(0).NumCells());
            Assert.AreEqual(3, outerTable.GetRow(1).NumCells());
            Assert.AreEqual(3, outerTable.GetRow(2).NumCells());
        }

    }

}
