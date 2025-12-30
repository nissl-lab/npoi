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
    using System.IO;
    using NUnit.Framework;using NUnit.Framework.Legacy;
    using NPOI;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Crypto;
    using NPOI.Util;

    /**
     * Tests for {@link RecordFactoryInputStream}
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestRecordFactoryInputStream
    {
        [TearDown]
        public void ResetPassword()
        {
            Biff8EncryptionKey.CurrentUserPassword = (null);
        }
        /**
         * Hex dump of a BOF record and most of a FILEPASS record.
         * A 16 byte saltHash should be Added to complete the second record
         */
        private static String COMMON_HEX_DATA = ""
            // BOF
            + "09 08 10 00"
            + "00 06  05 00  D3 10  CC 07  01 00 00 00  00 06 00 00"
            // FILEPASS
            + "2F 00 36 00"
            + "01 00  01 00  01 00"
            + "BAADF00D BAADF00D BAADF00D BAADF00D" // docId
            + "DEADBEEF DEADBEEF DEADBEEF DEADBEEF" // saltData
            ;

        /**
         * Hex dump of a sample WINDOW1 record
         */
        private static String SAMPLE_WINDOW1 = "3D 00 12 00"
            + "00 00 00 00 40 38 55 23 38 00 00 00 00 00 01 00 58 02";

        /**
         * Makes sure that a default password mismatch condition is represented with {@link EncryptedDocumentException}
         */
        [Test]
        [Ignore("not implemented")]
        public void TestDefaultPassword()
        {
            // This encodng depends on docId, password and stream position
            String SAMPLE_WINDOW1_ENCR1 = "3D 00 12 00"
                   + "C4, 9B, 02, 50, 86, E0, DF, 34, FB, 57, 0E, 8C, CE, 25, 45, E3, 80, 01";

            byte[] dataWrongDefault = HexRead.ReadFromString(""
                    + COMMON_HEX_DATA
                    + "00000000 00000000 00000000 00000001"
                    + SAMPLE_WINDOW1_ENCR1
            );

            RecordFactoryInputStream rfis;
            try
            {
                rfis = CreateRFIS(dataWrongDefault);
                throw new AssertionException("Expected password mismatch error");
            }
            catch (EncryptedDocumentException e)
            {
                // expected during successful Test
                if (!e.Message.Equals("Default password is invalid for docId/saltData/saltHash"))
                {
                    throw e;
                }
            }

            byte[] dataCorrectDefault = HexRead.ReadFromString(""
                    + COMMON_HEX_DATA
                    + "137BEF04 969A200B 306329DE 52254005" // correct saltHash for default password (and docId/saltHash)
                    + SAMPLE_WINDOW1_ENCR1
            );

            rfis = CreateRFIS(dataCorrectDefault);

            ConfirmReadInitialRecords(rfis);
        }

        /**
         * Makes sure that an incorrect user supplied password condition is represented with {@link EncryptedDocumentException}
         */
        [Test]
        [Ignore("not implemented")]
        public void TestSuppliedPassword()
        {
            // This encodng depends on docId, password and stream position
            String SAMPLE_WINDOW1_ENCR2 = "3D 00 12 00"
                   + "45, B9, 90, FE, B6, C6, EC, 73, EE, 3F, 52, 45, 97, DB, E3, C1, D6, FE";

            byte[] dataWrongDefault = HexRead.ReadFromString(""
                    + COMMON_HEX_DATA
                    + "00000000 00000000 00000000 00000000"
                    + SAMPLE_WINDOW1_ENCR2
            );


            Biff8EncryptionKey.CurrentUserPassword = (/*setter*/"passw0rd");

            RecordFactoryInputStream rfis;
            try
            {
                rfis = CreateRFIS(dataWrongDefault);
                throw new AssertionException("Expected password mismatch error");
            }
            catch (EncryptedDocumentException e)
            {
                // expected during successful Test
                if (!e.Message.Equals("Supplied password is invalid for docId/saltData/saltHash"))
                {
                    throw e;
                }
            }

            byte[] dataCorrectDefault = HexRead.ReadFromString(""
                    + COMMON_HEX_DATA
                    + "C728659A C38E35E0 568A338F C3FC9D70" // correct saltHash for supplied password (and docId/saltHash)
                    + SAMPLE_WINDOW1_ENCR2
            );

            rfis = CreateRFIS(dataCorrectDefault);
            Biff8EncryptionKey.CurrentUserPassword = (/*setter*/null);

            ConfirmReadInitialRecords(rfis);
        }

        /**
         * Makes sure the record stream starts with {@link BOFRecord} and then {@link WindowOneRecord}
         * The second record is Gets decrypted so this method also Checks its content.
         */
        private void ConfirmReadInitialRecords(RecordFactoryInputStream rfis)
        {
            ClassicAssert.AreEqual(typeof(BOFRecord), rfis.NextRecord().GetType());
            WindowOneRecord rec1 = (WindowOneRecord)rfis.NextRecord();
            ClassicAssert.IsTrue(Arrays.Equals(HexRead.ReadFromString(SAMPLE_WINDOW1), rec1.Serialize()));
        }

        private static RecordFactoryInputStream CreateRFIS(byte[] data)
        {
            return new RecordFactoryInputStream(new ByteArrayInputStream(data), true);
        }
    }

}