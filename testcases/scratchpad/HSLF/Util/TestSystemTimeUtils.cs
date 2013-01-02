using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.Util;
using NPOI.HSLF.Util;

namespace TestCases.HSLF.Util
{
    /// <summary>
    /// TestSystemTimeUtils 的摘要说明
    /// </summary>
    [TestClass]
    public class TestSystemTimeUtils
    {
        // From real files
        private byte[] data_a;
        private byte[] data_b;
        public TestSystemTimeUtils()
        {
            unchecked
            {
                data_a = new byte[] {
		(byte)(0xD6-256), 07, 01, 00,
		02, 00, 0x18, 00, 0x0A, 00, 0x1A, 00,
		0x0F, 00, (byte)(0xCD-256), 00
	};
                data_b = new byte[] {
		00, 00, (byte)(0xE1-256), 0x2E, 0x1C, 00, 00, 00,
		01, 00, 00, 00, (byte)(0xD6-256), 0x07, 01, 00,
		02, 00, 0x18, 00, 0x15, 00, 0x19, 00, 03,
		00, (byte)(0xD5-256), 02, 0x0A, 00, 00, 00,
		0x0A, 00, 00, 00
	};
            }
        }
        private SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss.SSS");
        [TestMethod]
        public void TestGetDateA()
        {
            DateTime date = SystemTimeUtils.GetDate(data_a);

            // Is 2006-01-24 (2nd day of week) 10:26:15.205
            DateTime exp = DateTime.Parse("2006-01-24 10:26:15.205");
            Assert.AreEqual(exp.TimeOfDay, date.TimeOfDay);
            Assert.AreEqual(exp, date);
        }
        [TestMethod]
        public void TestGetDateB()
        {
            DateTime date = SystemTimeUtils.GetDate(data_b, 8 + 4);

            // Is 2006-01-24 (2nd day of week) 21:25:03.725
            DateTime exp = DateTime.Parse("2006-01-24 21:25:03.725");
            Assert.AreEqual(exp.TimeOfDay, date.TimeOfDay);
            Assert.AreEqual(exp, date);
        }
        [TestMethod]
        public void TestWriteDateA()
        {
            byte[] out_a = new byte[data_a.Length];
            DateTime date = DateTime.Parse("2006-01-24 10:26:15.205");
            SystemTimeUtils.StoreDate(date, out_a);

            for (int i = 0; i < out_a.Length; i++)
            {
                Assert.AreEqual(data_a[i], out_a[i]);
            }
        }
        [TestMethod]
        public void TestWriteDateB()
        {
            byte[] out_b = new byte[data_b.Length];
            // Copy over start and end, ignoring the 16 byte date field in the middle
            Array.Copy(data_b, 0, out_b, 0, 12);
            Array.Copy(data_b, 12 + 16, out_b, 12 + 16, data_b.Length - 12 - 16);

            DateTime date = DateTime.Parse("2006-01-24 21:25:03.725");
            SystemTimeUtils.StoreDate(date, out_b, 12);

            for (int i = 0; i < out_b.Length; i++)
            {
                Assert.AreEqual(data_b[i], out_b[i]);
            }
        }
    }
}
