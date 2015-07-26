
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

namespace TestCases.DDF
{

    using System;
    using System.Text;
    using System.Collections.Generic;
    using System.IO;

    using NUnit.Framework;
    using NPOI.DDF;
    using NPOI.Util;
    [TestFixture]
    public class TestEscherOptRecord
    {
        [Test]
        public void TestFillFields()
        {
            CheckFillFieldsSimple();
            CheckFillFieldsComplex();
        }

        private void CheckFillFieldsComplex()
        {
            String dataStr = "33 00 " +
                    "0B F0 " +
                    "14 00 00 00 " +
                    "BF 00 01 00 00 00 " +
                    "01 80 02 00 00 00 " +
                    "BF 00 01 00 00 00 " +
                    "01 02";

            EscherOptRecord r = new EscherOptRecord();
            r.FillFields(HexRead.ReadFromString(dataStr), new DefaultEscherRecordFactory());
            Assert.AreEqual((short)0x0033, r.Options);
            Assert.AreEqual(unchecked((short)0xF00B), r.RecordId);
            Assert.AreEqual(3, r.EscherProperties.Count);
            EscherBoolProperty prop1 = new EscherBoolProperty(EscherProperties.TEXT__SIZE_TEXT_TO_FIT_SHAPE, 1);
            EscherComplexProperty prop2 = new EscherComplexProperty((short)1, false, new byte[] { 0x01, 0x02 });
            EscherBoolProperty prop3 = new EscherBoolProperty(EscherProperties.TEXT__SIZE_TEXT_TO_FIT_SHAPE, 1);
            Assert.AreEqual(prop1, r.GetEscherProperty(0));
            Assert.AreEqual(prop2, r.GetEscherProperty(1));
            Assert.AreEqual(prop3, r.GetEscherProperty(2));

        }

        private void CheckFillFieldsSimple()
        {
            String dataStr = "33 00 " + // options
                            "0B F0 " + // recordid
                            "12 00 00 00 " + // remaining bytes
                            "BF 00 08 00 08 00 " +
                            "81 01 09 00 00 08 " +
                            "C0 01 40 00 00 08";

            EscherOptRecord r = new EscherOptRecord();
            r.FillFields(HexRead.ReadFromString(dataStr), new DefaultEscherRecordFactory());
            Assert.AreEqual((short)0x0033, r.Options);
            Assert.AreEqual(unchecked((short)0xF00B), r.RecordId);
            Assert.AreEqual(3, r.EscherProperties.Count);
            EscherBoolProperty prop1 = new EscherBoolProperty(EscherProperties.TEXT__SIZE_TEXT_TO_FIT_SHAPE, 524296);
            EscherRGBProperty prop2 = new EscherRGBProperty(EscherProperties.FILL__FILLCOLOR, 0x08000009);
            EscherRGBProperty prop3 = new EscherRGBProperty(EscherProperties.LINESTYLE__COLOR, 0x08000040);
            Assert.AreEqual(prop1, r.GetEscherProperty(0));
            Assert.AreEqual(prop2, r.GetEscherProperty(1));
            Assert.AreEqual(prop3, r.GetEscherProperty(2));
        }
        [Test]
        public void TestSerialize()
        {
            CheckSerializeSimple();
            CheckSerializeComplex();
        }

        private void CheckSerializeComplex()
        {
            //Complex escher record
            EscherOptRecord r = new EscherOptRecord();
            r.Options=(short)0x0033;
            r.RecordId=unchecked((short)0xF00B);
            EscherBoolProperty prop1 = new EscherBoolProperty(EscherProperties.TEXT__SIZE_TEXT_TO_FIT_SHAPE, 1);
            EscherComplexProperty prop2 = new EscherComplexProperty((short)1, false, new byte[] { 0x01, 0x02 });
            EscherBoolProperty prop3 = new EscherBoolProperty(EscherProperties.TEXT__SIZE_TEXT_TO_FIT_SHAPE, 1);
            r.AddEscherProperty(prop1);
            r.AddEscherProperty(prop2);
            r.AddEscherProperty(prop3);

            byte[] data = new byte[28];
            int bytesWritten = r.Serialize(0, data);
            Assert.AreEqual(28, bytesWritten);
            String dataStr = "[33, 00, " +
                    "0B, F0, " +
                    "14, 00, 00, 00, " +
                    "BF, 00, 01, 00, 00, 00, " +
                    "01, 80, 02, 00, 00, 00, " +
                    "BF, 00, 01, 00, 00, 00, " +
                    "01, 02]";
            Assert.AreEqual(dataStr, HexDump.ToHex(data));

        }

        private void CheckSerializeSimple()
        {
            EscherOptRecord r = new EscherOptRecord();
            r.Options=(short)0x0033;
            r.RecordId=unchecked((short)0xF00B);
            EscherBoolProperty prop1 = new EscherBoolProperty(EscherProperties.TEXT__SIZE_TEXT_TO_FIT_SHAPE, 1);
            EscherRGBProperty prop2 = new EscherRGBProperty(EscherProperties.FILL__FILLCOLOR, 0x08000009);
            EscherRGBProperty prop3 = new EscherRGBProperty(EscherProperties.LINESTYLE__COLOR, 0x08000040);
            r.AddEscherProperty(prop1);
            r.AddEscherProperty(prop2);
            r.AddEscherProperty(prop3);

            byte[] data = new byte[26];
            int bytesWritten = r.Serialize(0, data);
            String dataStr = "[33, 00, " +
                    "0B, F0, " +
                    "12, 00, 00, 00, " +
                    "BF, 00, 01, 00, 00, 00, " +
                    "81, 01, 09, 00, 00, 08, " +
                    "C0, 01, 40, 00, 00, 08]";
            Assert.AreEqual(dataStr, HexDump.ToHex(data));
            Assert.AreEqual(26, bytesWritten);
        }
        [Test]
        public void TestToString()
        {
            String nl = Environment.NewLine;
            EscherOptRecord r = new EscherOptRecord();
            // don't try to shoot in foot, please -- vlsergey
            // r.setOptions((short)0x000F);
            r.RecordId=EscherOptRecord.RECORD_ID;
            EscherProperty prop1 = new EscherBoolProperty((short)1, 1);
            r.AddEscherProperty(prop1);
            String expected = "EscherOptRecord:" + nl +
                    "  isContainer: False" + nl +
                    "  version: 0x0003" + nl +
                    "  instance: 0x0001" + nl +
                    "  recordId: 0x" + HexDump.ToHex(EscherOptRecord.RECORD_ID) + nl +
                    "  numchildren: 0" + nl +
                    "  properties:" + nl +
                    "    propNum: 1, RAW: 0x0001, propName: unknown, complex: False, blipId: False, value: 1 (0x00000001)" + nl;
            Assert.AreEqual(expected, r.ToString());
        }

        /**
         * Test serialisation of a particually complex example 
         * This Test is currently broken!
         */
        [Test]
        public void TestComplexSerialise()
        {
            byte[] data = new byte[] {
    		0x53, 0x01, 0x0B, unchecked((byte)(0xF0-256)), unchecked((byte)(0x9C-256)), 0x01, 0x00, 0x00, 
    		// Simple data follows
    		0x42, 0x01,	0x49, 0x00, 0x00, 0x00,          // SP @ 8
    		0x43, 0x01, unchecked((byte)(0x85-256)), 0x00, 0x00, 0x00,      // SP @ 14
    		0x44, 0x01, 0x04, 0x00, 0x00, 0x00,          // SP @ 20
    		0x45, unchecked((byte)(0xC1-256)), unchecked((byte)(0x88-256)), 0x00, 0x00, 0x00,  // SP @ 26
    		0x46, unchecked((byte)(0xC1-256)), unchecked((byte)(0x90-256)), 0x00, 0x00, 0x00,  // SP @ 32
    		0x7F, 0x01, 0x01, 0x00, 0x01, 0x00, 
    		unchecked((byte)(0x80-256)), 0x01, 0x00, 0x00, 0x00, 0x00, 
    		unchecked((byte)(0x81-256)), 0x01, 0x02, 0x00, 0x00, 0x08, 
    		unchecked((byte)(0xBF-256)), 0x01,	0x10, 0x00, 0x10, 0x00, 
    		unchecked((byte)(0xC0-256)), 0x01, 0x01, 0x00, 0x00, 0x08,      // SP 10 
    		unchecked((byte)(0xC1-256)), 0x01, 0x00, 0x00, 0x01, 0x00,
    		unchecked((byte)(0xC4-256)), 0x01, 0x00, 0x00, 0x00, 0x00, 
    		unchecked((byte)(0xCB-256)), 0x01, 0x38, 0x63, 0x00, 0x00, 
    		unchecked((byte)(0xCD-256)), 0x01, 0x00, 0x00,	0x00, 0x00, 
    		unchecked((byte)(0xCE-256)), 0x01, 0x00, 0x00, 0x00, 0x00,      // SP 15
    		unchecked((byte)(0xD0-256)), 0x01, 0x00, 0x00, 0x00, 0x00, 
    		unchecked((byte)(0xD1-256)), 0x01, 0x00, 0x00, 0x00, 0x00, 
    		unchecked((byte)(0xD7-256)), 0x01, 0x00, 0x00, 0x00, 0x00, 
    		unchecked((byte)(0xFF-256)), 0x01, 0x18, 0x00, 0x18, 0x00,
    		0x01, 0x02, 0x02, 0x00, 0x00, 0x08, 
    		0x3F, 0x02, 0x00, 0x00,	0x02, 0x00,          // SP 21
    		
    		// Complex data follows
    		
    		// Complex data for Array #325
    		// Array header
    		0x22, 0x00, 0x22, 0x00, unchecked((byte)(0xF0-256)), unchecked((byte)(0xFF-256)), 
    		// Array data
    		0x18, 0x00, 0x28, 0x00, 0x04, 0x00, 0x34, 
    		0x00, 0x04, 0x00, 0x28, 0x00, 0x04, 0x00,
    		0x1C, 0x00, 0x04, 0x00, 0x10, 0x00, 0x04, 0x00, 0x04, 0x00, 0x10, 
    		0x00, 0x00, 0x00, 0x1C, 0x00,
    		0x04, 0x00, 0x28, 0x00, 0x10, 0x00, 0x34, 0x00, 0x18, 0x00, 0x3C, 
    		0x00, 0x24, 0x00, 0x44, 0x00,
    		0x30, 0x00, 0x48, 0x00, 0x3C, 0x00, 0x44, 0x00, 0x48, 0x00, 0x3C, 
    		0x00, 0x54, 0x00, 0x38, 0x00,
    		0x60, 0x00, 0x2C, 0x00, 0x70, 0x00, 0x20, 0x00, 0x78, 0x00, 
    		0x14, 0x00, unchecked((byte)(0x80-256)), 0x00, 0x08, 0x00,
    		unchecked((byte)(0x84-256)), 0x00, 0x04, 0x00, 0x78, 0x00, 0x04, 0x00, 0x6C, 0x00, 
    		0x04, 0x00, 0x60, 0x00, 0x04, 0x00,
    		0x54, 0x00, 0x08, 0x00, 0x48, 0x00, 0x0C, 0x00, 0x3C, 0x00, 0x0C, 
    		0x00, 0x30, 0x00, 0x08, 0x00,
    		0x3C, 0x00, 0x08, 0x00, 0x48, 0x00, 0x08, 0x00, 0x54, 0x00, 0x00, 
    		0x00, 0x48, 0x00, 0x00, 0x00,
    		0x3C, 0x00, 0x00, 0x00, 0x30, 0x00, 0x04, 0x00, 0x24, 0x00,
    		// Complex data for Array #326
    		// Array header
    		0x45, 0x00, 0x48, 0x00, 0x02, 0x00, 
    		// Array data
    		0x00, 0x40, 0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 
    		0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)),
    		0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 
    		0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)),
    		0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 
    		0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)),
    		0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 
    		0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)),
    		0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 
    		0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)),
    		0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 
    		0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)),
    		0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 0x00,unchecked((byte)(0xB0-256)), 0x01, 0x00, 
    		0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)),
    		0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 
    		0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)),
    		0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)), 0x01, 0x00, 0x00, unchecked((byte)(0xB0-256)), 0x00, unchecked((byte)(0x80-256))
    	};

            // Create the record
            EscherOptRecord r = new EscherOptRecord();
            int Filled = r.FillFields(data, new DefaultEscherRecordFactory());

            // Check it's the right Length
            Assert.AreEqual(data.Length, Filled);
            Assert.AreEqual(data.Length, r.RecordSize);

            // Serialise it
            byte[] dest = new byte[data.Length];
            int written = r.Serialize(0, dest);

            // Check it serialised it back to the same data
            Assert.AreEqual(data.Length, written);
            for (int i = 0; i < data.Length; i++)
            {
                Assert.AreEqual(data[i], dest[i]);
            }
        }

        /**
         * Test Read/Write against an OPT record from a real ppt file.
         * In PowerPoint is is legal to have array properties with empty complex part.
         * In Glen's original implementation complex part is always 6 bytes which resulted
         * in +6 extra bytes when writing back out. As the result the ppt becomes unReadable.
         *
         * Make sure we Write back the original empty complex part.
         *
         * See Bug 41946 for details.
         */
        [Test]
        public void Test41946()
        {
            String dataStr1 =
                    "03 08 0B F0 00 03 00 00 81 00 30 65 01 00 82 00 98 B2 00 00 83 00 30 65 01 " +
                    "00 84 00 98 B2 00 00 85 00 00 00 00 00 87 00 01 00 00 00 88 00 00 00 00 00 " +
                    "89 00 00 00 00 00 BF 00 00 00 0F 00 0C 01 F4 00 00 10 0D 01 00 00 00 20 0E " +
                    "01 00 00 00 20 80 01 00 00 00 00 81 01 04 00 00 08 82 01 00 00 01 00 83 01 " +
                    "00 00 00 08 84 01 00 00 01 00 85 01 00 00 00 20 86 41 00 00 00 00 87 C1 00 " +
                    "00 00 00 88 01 00 00 00 00 89 01 00 00 00 00 8A 01 00 00 00 00 8B 01 00 00 " +
                    "00 00 8C 01 00 00 00 00 8D 01 00 00 00 00 8E 01 00 00 00 00 8F 01 00 00 00 " +
                    "00 90 01 00 00 00 00 91 01 00 00 00 00 92 01 00 00 00 00 93 01 00 00 00 00 " +
                    "94 01 00 00 00 00 95 01 00 00 00 00 96 01 00 00 00 00 97 C1 00 00 00 00 98 " +
                    "01 00 00 00 00 99 01 00 00 00 00 9A 01 00 00 00 00 9B 01 00 00 00 00 9C 01 " +
                    "03 00 00 40 BF 01 0C 00 1E 00 C0 01 01 00 00 08 C1 01 00 00 01 00 C2 01 FF " +
                    "FF FF 00 C3 01 00 00 00 20 C4 01 00 00 00 00 C5 41 00 00 00 00 C6 C1 00 00 " +
                    "00 00 C7 01 00 00 00 00 C8 01 00 00 00 00 C9 01 00 00 00 00 CA 01 00 00 00 " +
                    "00 CB 01 35 25 00 00 CC 01 00 00 08 00 CD 01 00 00 00 00 CE 01 00 00 00 00 " +
                    "CF C1 00 00 00 00 D7 01 02 00 00 00 FF 01 06 00 0E 00 00 02 00 00 00 00 01 " +
                    "02 02 00 00 08 02 02 CB CB CB 00 03 02 00 00 00 20 04 02 00 00 01 00 05 02 " +
                    "38 63 00 00 06 02 38 63 00 00 07 02 00 00 00 00 08 02 00 00 00 00 09 02 00 " +
                    "00 01 00 0A 02 00 00 00 00 0B 02 00 00 00 00 0C 02 00 00 01 00 0D 02 00 00 " +
                    "00 00 0E 02 00 00 00 00 0F 02 00 01 00 00 10 02 00 00 00 00 11 02 00 00 00 " +
                    "00 3F 02 00 00 03 00 80 02 00 00 00 00 81 02 00 00 01 00 82 02 05 00 00 00 " +
                    "83 02 9C 31 00 00 84 02 00 00 00 00 85 02 F0 F9 06 00 86 02 00 00 00 00 87 " +
                    "02 F7 00 00 10 88 02 00 00 00 20 BF 02 01 00 0F 00 C0 02 00 00 00 00 C1 02 " +
                    "00 00 00 00 C2 02 64 00 00 00 C3 02 00 00 00 00 C4 02 00 00 00 00 C5 02 00 " +
                    "00 00 00 C6 02 00 00 00 00 C7 02 00 00 00 00 C8 02 00 00 00 00 C9 02 00 00 " +
                    "00 00 CA 02 30 75 00 00 CB 02 D0 12 13 00 CC 02 30 ED EC FF CD 02 40 54 89 " +
                    "00 CE 02 00 80 00 00 CF 02 00 80 FF FF D0 02 00 00 79 FF D1 02 32 00 00 00 " +
                    "D2 02 20 4E 00 00 D3 02 50 C3 00 00 D4 02 00 00 00 00 D5 02 10 27 00 00 D6 " +
                    "02 70 94 00 00 D7 02 B0 3C FF FF D8 02 00 00 00 00 D9 02 10 27 00 00 DA 02 " +
                    "70 94 00 00 FF 02 16 00 1F 00 04 03 01 00 00 00 41 03 A8 29 01 00 42 03 00 " +
                    "00 00 00 43 03 03 00 00 00 44 03 7C BE 01 00 45 03 00 00 00 00 7F 03 00 00 " +
                    "0F 00 84 03 7C BE 01 00 85 03 00 00 00 00 86 03 7C BE 01 00 87 03 00 00 00 " +
                    "00";

            EscherOptRecord r = new EscherOptRecord();
            byte[] data = HexRead.ReadFromString(dataStr1);
            r.FillFields(data, 0, new DefaultEscherRecordFactory());
            Assert.AreEqual(unchecked((short)0xF00B), r.RecordId);

            byte[] data1 = r.Serialize();
            EscherOptRecord opt2 = new EscherOptRecord();
            opt2.FillFields(data1, new DefaultEscherRecordFactory());

            byte[] data2 = opt2.Serialize();
            Assert.IsTrue(Arrays.Equals(data1, data2));
        }

        /**
         * Test that EscherOptRecord can properly Read/Write array properties
         * with empty complex part.
         */
        [Test]
        public void TestEmptyArrayProperty()
        {
            EscherOptRecord r = new EscherOptRecord();
            EscherArrayProperty p = new EscherArrayProperty(unchecked((short)(EscherProperties.FILL__SHADECOLORS + 0x8000)), new byte[0]);
            Assert.AreEqual(0, p.NumberOfElementsInArray);
            r.AddEscherProperty(p);

            byte[] data1 = r.Serialize();
            EscherOptRecord opt2 = new EscherOptRecord();
            opt2.FillFields(data1, new DefaultEscherRecordFactory());
            p = (EscherArrayProperty)opt2.EscherProperties[0];
            Assert.AreEqual(0, p.NumberOfElementsInArray);

            byte[] data2 = opt2.Serialize();
            Assert.IsTrue(Arrays.Equals(data1, data2));
        }
    }
}