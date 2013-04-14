using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using NPOI.HSSF.Record;
using NPOI.Util;
using NPOI.SS.UserModel;

namespace TestCases.HSSF.Record
{
    [TestFixture]
    public class TestFontRecord
    {
        private const int SID = 0x31;
        private static byte[] data = {
                    unchecked((byte)(0xC8-256)), 00,       // font height = xc8
                    00, 00,             // attrs = 0
                    unchecked((byte)(0xFF-256)), 0x7F,     // colour palette = x7fff
                    unchecked((byte)(0x90-256)), 0x01,     // bold weight = x190
                    00, 00,  // supersubscript
                    00, 00,  // underline, family
                    00, 00,  // charset, padding
                    05, 00,  // name length, unicode flag
                    0x41, 0x72, 0x69, 0x61, 0x6C, // Arial, as unicode

            };
        [Test]
        public void TestLoad()
        {

            FontRecord record = new FontRecord(TestcaseRecordInputStream.Create(0x31, data));
            Assert.AreEqual(0xc8, record.FontHeight);
            Assert.AreEqual(0x00, record.Attributes);
            Assert.IsFalse(record.IsItalic);
            Assert.IsFalse(record.IsStrikeout);
            Assert.IsFalse(record.IsMacoutlined);
            Assert.IsFalse(record.IsMacshadowed);
            Assert.AreEqual(0x7fff, record.ColorPaletteIndex);
            Assert.AreEqual(0x190, record.BoldWeight);
            Assert.AreEqual(FontSuperScript.None, record.SuperSubScript);
            Assert.AreEqual(FontUnderlineType.None, record.Underline);
            Assert.AreEqual(0x00, record.Family);
            Assert.AreEqual(0x00, record.Charset);
            Assert.AreEqual("Arial", record.FontName);

            Assert.AreEqual(21 + 4, record.RecordSize);
        }
        [Test]
        public void TestStore()
        {
            //      .fontheight      = c8
            //      .attributes      = 0
            //           .italic     = false
            //           .strikout   = false
            //           .macoutlined= false
            //           .macshadowed= false
            //      .colorpalette    = 7fff
            //      .boldweight      = 190
            //      .supersubscript  = 0
            //      .underline       = 0
            //      .family          = 0
            //      .charset         = 0
            //      .namelength      = 5
            //      .fontname        = Arial

            FontRecord record = new FontRecord();
            record.FontHeight = ((short)0xc8);
            record.Attributes=((short)0);
            record.ColorPaletteIndex=((short)0x7fff);
            record.BoldWeight=((short)0x190);
            record.SuperSubScript=((short)0);
            record.Underline=((byte)0);
            record.Family=((byte)0);
            record.Charset=((byte)0);
            record.FontName = ("Arial");

            byte[] recordBytes = record.Serialize();
            TestcaseRecordInputStream.ConfirmRecordEncoding(0x31, data, recordBytes);
        }
        [Test]
        public void TestCloneOnto()
        {
            FontRecord base1 = new FontRecord(TestcaseRecordInputStream.Create(0x31, data));

            FontRecord other = new FontRecord();
            other.CloneStyleFrom(base1);

            byte[] recordBytes = other.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
        [Test]
        public void TestSameProperties()
        {
            FontRecord f1 = new FontRecord(TestcaseRecordInputStream.Create(0x31, data));
            FontRecord f2 = new FontRecord(TestcaseRecordInputStream.Create(0x31, data));

            Assert.IsTrue(f1.SameProperties(f2));

            f2.FontName = ("Arial2");
            Assert.IsFalse(f1.SameProperties(f2));
            f2.FontName = ("Arial");
            Assert.IsTrue(f1.SameProperties(f2));

            f2.FontHeight = ((short)11);
            Assert.IsFalse(f1.SameProperties(f2));
            f2.FontHeight = ((short)0xc8);
            Assert.IsTrue(f1.SameProperties(f2));
        }

        /**
         * Bugzilla 47250 suggests that the unicode options byte should be present even when the name
         * length is zero.  The OOO documentation seems to agree with this and POI had no test data
         * samples to say otherwise.
         */
        [Test]
        public void TestEmptyName_bug47250()
        {
            byte[] emptyNameData = HexRead.ReadFromString(
                    "C8 00 00 00 FF 7F 90 01 00 00 00 00 00 00 "
                    + "00" // zero length
                    + "00" // unicode options byte
                    );

            RecordInputStream in1 = TestcaseRecordInputStream.Create(SID, emptyNameData);
            FontRecord fr = new FontRecord(in1);
            if (in1.Available() == 1)
            {
                throw new AssertionException("Identified bug 47250");
            }
            Assert.AreEqual(0, in1.Available());

            Assert.AreEqual(0, fr.FontName.Length);
            byte[] recordBytes = fr.Serialize();
            TestcaseRecordInputStream.ConfirmRecordEncoding(SID, emptyNameData, recordBytes);
        }
    }
}
