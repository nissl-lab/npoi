using NUnit.Framework;
using NPOI.HSSF.Record;
using NPOI.HSSF.Record.AutoFilter;
using NPOI.Util;
using NUnit.Framework.Legacy;

namespace TestCases.HSSF.Record
{
    [TestFixture]
    public class TestAutoFilterRecord
    {
        private readonly byte[] recordData =
        [
            0x01, 0x00, // iEntry
            0x01, 0x00, // grBit

            06, // doper1 string
            0x00, // sign
            0x00, 0x00, 0x00, 0x00, // reserved
            0x02, // cch (length of string)
            0x00, // reserved
            0x00, 0x00, // reserved,

            06, // doper2 string
            0x00, // sign
            0x00, 0x00, 0x00, 0x00, // reserved
            0x02, // cch (length of string)
            0x00, // reserved
            0x00, 0x00, // reserved,

            0x00, // non-multibyte
            (byte)'a',
            (byte)'b',

            0x01, // multibyte
            (byte)'a', 0x00,
            (byte)'b', 0x00,
        ];

        [Test]
        public void TestRead()
        {
            AutoFilterRecord record = new AutoFilterRecord(TestcaseRecordInputStream.Create(AutoFilterRecord.sid, recordData));

            ClassicAssert.AreEqual(AutoFilterRecord.sid, record.Sid);
            ClassicAssert.AreEqual(recordData.Length, record.RecordSize - 4);
            ClassicAssert.AreEqual(1, record.iEntry);
            ClassicAssert.AreEqual(DOPERType.String, record.Doper1.DataType);
            ClassicAssert.AreEqual(DOPERType.String, record.Doper2.DataType);
            ClassicAssert.AreEqual("ab", record.Doper1RGCH);
            ClassicAssert.AreEqual("ab", record.Doper2RGCH);
        }

        [Test]
        public void TestWrite()
        {
            AutoFilterRecord record = new AutoFilterRecord(TestcaseRecordInputStream.Create(AutoFilterRecord.sid, recordData))
            {
                Doper1RGCH = "Multibyte - абвгд", // contains Cyrillic letters
                Doper2RGCH = "Non multibyte"
            };

            byte[] ser = record.Serialize();
            ClassicAssert.AreEqual(77, ser.Length);

            record = new AutoFilterRecord(TestcaseRecordInputStream.Create(ser));
            ClassicAssert.AreEqual("Multibyte - абвгд", record.Doper1RGCH);
            ClassicAssert.AreEqual("Non multibyte", record.Doper2RGCH);
        }

        [Test]
        public void TestClone()
        {
            AutoFilterRecord record = new AutoFilterRecord(TestcaseRecordInputStream.Create(AutoFilterRecord.sid, recordData))
            {
                Doper1RGCH = "Multibyte - абвгд", // contains Cyrillic letters
                Doper2RGCH = "Non multibyte",
                iEntry = (/*setter*/(short)3)
            };
            byte[] src = record.Serialize();

            AutoFilterRecord Cloned = (AutoFilterRecord)record.Clone();
            ClassicAssert.AreEqual(3, record.iEntry);
            ClassicAssert.AreEqual("Multibyte - абвгд", Cloned.Doper1RGCH);
            ClassicAssert.AreEqual("Non multibyte", Cloned.Doper2RGCH);
            byte[] cln = Cloned.Serialize();

            ClassicAssert.AreEqual(record.RecordSize, Cloned.RecordSize);
            ClassicAssert.IsTrue(Arrays.Equals(src, cln));
        }

        [Test]
        public void TestMultibyteDataSize()
        {
            AutoFilterRecord record = new AutoFilterRecord();
            record.Doper1RGCH = "日本語テスト"; // Japanese multibyte characters
            record.Doper2RGCH = "中文测试"; // Chinese multibyte characters

            int expectedSize = 2 + 2 + 10 + 10;
            expectedSize += 1 + (6 * 2); // doper1: multibyte flag + 6 chars * 2 bytes
            expectedSize += 1 + (4 * 2); // doper2: multibyte flag + 4 chars * 2 bytes

            ClassicAssert.AreEqual(expectedSize, record.RecordSize - 4);

            byte[] ser = record.Serialize();
            ClassicAssert.AreEqual(4 + expectedSize, ser.Length);

            AutoFilterRecord deserialized = new AutoFilterRecord(TestcaseRecordInputStream.Create(ser));
            ClassicAssert.AreEqual("日本語テスト", deserialized.Doper1RGCH);
            ClassicAssert.AreEqual("中文测试", deserialized.Doper2RGCH);
        }

        [Test]
        public void TestMixedMultibyteDataSize()
        {
            AutoFilterRecord record = new AutoFilterRecord();
            record.Doper1RGCH = "test";
            record.Doper2RGCH = "データ";

            int baseSize = 2 + 2 + 10 + 10;
            int doper1Size = 1 + record.Doper1RGCH.Length * (StringUtil.HasMultibyte(record.Doper1RGCH) ? 2 : 1);
            int doper2Size = 1 + record.Doper2RGCH.Length * (StringUtil.HasMultibyte(record.Doper2RGCH) ? 2 : 1);

            int expectedSize = baseSize + doper1Size + doper2Size;
            ClassicAssert.AreEqual(expectedSize, record.RecordSize - 4);

            byte[] ser = record.Serialize();
            ClassicAssert.AreEqual(4 + expectedSize, ser.Length);

            AutoFilterRecord deserialized = new AutoFilterRecord(TestcaseRecordInputStream.Create(ser));
            ClassicAssert.AreEqual("test", deserialized.Doper1RGCH);
            ClassicAssert.AreEqual("データ", deserialized.Doper2RGCH);
        }
    }
}