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

namespace TestCases.HSSF.Record
{
    using System;

    using NUnit.Framework;
    using TestCases.HSSF.Record;
    using NPOI.HSSF.Record;
    /**
     */
    [TestFixture]
    public class TestExtendedFormatRecord
    {

        private static byte[] data = new byte[] {
			00, 00, // Font 0
			00, 00, // Format 0
			unchecked((byte)(0xF5 - 256)), unchecked((byte)(0xFF - 256)), // Cell opts ...
			0x20, 00, // Alignment 20
			00, 00, // Ident 0
			00, 00, // Border 0
			00, 00, // Palette 0
			00, 00, 00, 00, // ADTL Palette 0
			unchecked((byte)(0xC0 - 256)), 0x20 // Fill Palette 20c0
	};

        private static ExtendedFormatRecord CreateEFR()
        {
            return new ExtendedFormatRecord(TestcaseRecordInputStream.Create(0x00E0, data));
        }
        [Test]
        public void TestLoad()
        {
            ExtendedFormatRecord record = CreateEFR();
            Assert.AreEqual(0, record.FontIndex);
            Assert.AreEqual(0, record.FormatIndex);
            Assert.AreEqual(0xF5 - 256, record.CellOptions);
            Assert.AreEqual(0x20, record.AlignmentOptions);
            Assert.AreEqual(0, record.IndentionOptions);
            Assert.AreEqual(0, record.BorderOptions);
            Assert.AreEqual(0, record.PaletteOptions);
            Assert.AreEqual(0, record.AdtlPaletteOptions);
            Assert.AreEqual(0x20c0, record.FillPaletteOptions);

            Assert.AreEqual(20 + 4, record.RecordSize);
        }

        [Test]
        public void TestStore()
        {
            //    .fontindex       = 0
            //    .formatindex     = 0
            //    .celloptions     = fffffff5
            //          .Islocked  = true
            //          .Ishidden  = false
            //          .Recordtype= 1
            //          .parentidx = fff
            //    .alignmentoptions= 20
            //          .alignment = 0
            //          .wraptext  = false
            //          .valignment= 2
            //          .justlast  = 0
            //          .rotation  = 0
            //    .indentionoptions= 0
            //          .indent    = 0
            //          .shrinktoft= false
            //          .mergecells= false
            //          .Readngordr= 0
            //          .formatflag= false
            //          .fontflag  = false
            //          .prntalgnmt= false
            //          .borderflag= false
            //          .paternflag= false
            //          .celloption= false
            //    .borderoptns     = 0
            //          .lftln     = 0
            //          .rgtln     = 0
            //          .Topln     = 0
            //          .btmln     = 0
            //    .paleteoptns     = 0
            //          .leftborder= 0
            //          .rghtborder= 0
            //          .diag      = 0
            //    .paleteoptn2     = 0
            //          .Topborder = 0
            //          .botmborder= 0
            //          .adtldiag  = 0
            //          .diaglnstyl= 0
            //          .Fillpattrn= 0
            //    .Fillpaloptn     = 20c0
            //          .foreground= 40
            //          .background= 41

            ExtendedFormatRecord record = new ExtendedFormatRecord();
            record.FontIndex = (/*setter*/(short)0);
            record.FormatIndex = (/*setter*/(short)0);

            record.IsLocked = (/*setter*/true);
            record.IsHidden = (/*setter*/false);
            record.XFType = (/*setter*/(short)1);
            record.ParentIndex = (/*setter*/(short)0xfff);

            record.VerticalAlignment = (/*setter*/(short)2);

            record.FillForeground = (/*setter*/(short)0x40);
            record.FillBackground = (/*setter*/(short)0x41);

            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
        [Test]
        public void TestCloneOnto()
        {
            ExtendedFormatRecord base1 = CreateEFR();

            ExtendedFormatRecord other = new ExtendedFormatRecord();
            other.CloneStyleFrom(base1);

            byte[] recordBytes = other.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }

        [Test]
        public void TestRotation()
        {
            ExtendedFormatRecord record = CreateEFR();
            Assert.AreEqual(0, record.Rotation);
            record.Rotation = ((short)1);
            Assert.AreEqual(1, record.Rotation);
            record.Rotation = ((short)89);
            Assert.AreEqual(89, record.Rotation);
            record.Rotation = ((short)90);
            Assert.AreEqual(90, record.Rotation);
            // internally values below zero are stored differently
            record.Rotation = ((short)-1);
            Assert.AreEqual(255, record.Rotation);
            record.Rotation = ((short)-89);
            Assert.AreEqual(-77, 90 - record.Rotation);
            record.Rotation = ((short)-90);
            Assert.AreEqual(-76, 90 - record.Rotation);
        }
    }

}