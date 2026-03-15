using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.XWPF.UserModel;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using System.Collections.Generic;

namespace TestCases.XWPF.UserModel
{
    [TestFixture]
    public class TestXWPFChart
    {
        [Test(Description = "Test method to check charts are null")]
        public void TestRead()
        {
            XWPFDocument sampleDoc = XWPFTestDataSamples.OpenSampleDocument("61745.docx");
            List<XWPFChart> charts = sampleDoc.GetCharts();
            ClassicAssert.IsNotNull(charts);
            ClassicAssert.AreEqual(2, charts.Count);
            ClassicAssert.IsNotNull(charts[0]);
            ClassicAssert.IsNotNull(charts[1]);
        }

        [Test(Description = "Test method to add chart title and check whether it's set")]
        public void TestChartTitle()
        {
            XWPFDocument sampleDoc = XWPFTestDataSamples.OpenSampleDocument("61745.docx");
            List<XWPFChart> charts = sampleDoc.GetCharts();
            XWPFChart chart=charts[0];
            CT_Chart ctChart = chart.GetCTChart();
            CT_Title title = ctChart.title;
            CT_Tx tx = title.AddNewTx();
            CT_TextBody rich = tx.AddNewRich();
            rich.AddNewBodyPr();
            rich.AddNewLstStyle();
            CT_TextParagraph p = rich.AddNewP();
            CT_RegularTextRun r = p.AddNewR();
            r.AddNewRPr();
            r.t = ("XWPF CHART");
            ClassicAssert.AreEqual("XWPF CHART", chart.GetCTChart().title.tx.rich.GetPArray(0).GetRArray(0).t);
        }

        [Test(Description = "Test method to check relationship")]
        public void TestChartRelation()
        {
            XWPFDocument sampleDoc = XWPFTestDataSamples.OpenSampleDocument("61745.docx");
            List<XWPFChart> charts = sampleDoc.GetCharts();
            XWPFChart chart=charts[0];
            ClassicAssert.AreEqual(XWPFRelation.CHART.ContentType, chart.GetPackagePart().ContentType);
            ClassicAssert.AreEqual("/word/document.xml", chart.GetParent().GetPackagePart().PartName);
            ClassicAssert.AreEqual("/word/charts/chart1.xml", chart.GetPackagePart().PartName);
        }
    }
}
