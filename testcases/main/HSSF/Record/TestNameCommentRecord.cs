using System;
using System.Text;
using System.Collections.Generic;

using NUnit.Framework;
using NPOI.Util;
using NPOI.HSSF.Record;

namespace TestCases.HSSF.Record
{
    /// <summary>
    /// Summary description for TestNameCommentRecord
    /// </summary>
    [TestFixture]
    public class TestNameCommentRecord
    {
        public TestNameCommentRecord()
        {
            //
            // TODO: Add constructor logic here
            //
        }


        [Test]
        public void TestReserialize()
        {
            byte[] data = HexRead
                    .ReadFromString(""
                            + "94 08 00 00 00 00 00 00 00 00 00 00 04 00 07 00 00 6E 61 6D 65 00 63 6F 6D 6D 65 6E 74]");
            RecordInputStream in1 = TestcaseRecordInputStream.Create(NameCommentRecord.sid, data);
            NameCommentRecord ncr = new NameCommentRecord(in1);
            Assert.AreEqual(0x0894, ncr.RecordType);
            Assert.AreEqual("name", ncr.NameText);
            Assert.AreEqual("comment", ncr.CommentText);
            byte[] data2 = ncr.Serialize();
            TestcaseRecordInputStream.ConfirmRecordEncoding(NameCommentRecord.sid, data, data2);
        }
    }
}
