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

namespace TestCases.HSSF.Record.Crypto
{
    using System;
    using System.Text;
    using NUnit.Framework;
    using NPOI.Util;
    using TestCases.Exceptions;
    using NPOI.HSSF.Record.Crypto;
    /**
     * Tests for {@link RC4}
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestRC4
    {
        [Test]
        public void TestSimple()
        {
            ConfirmRC4("Key", "Plaintext", "BBF316E8D940AF0AD3");
            ConfirmRC4("Wiki", "pedia", "1021BF0420");
            ConfirmRC4("Secret", "Attack at dawn", "45A01F645FC35B383552544B9BF5");

        }

        private static void ConfirmRC4(String k, String origText, String expEncrHex)
        {
            byte[] actEncr = Encoding.GetEncoding("GB2312").GetBytes(origText);
            new RC4(Encoding.GetEncoding("GB2312").GetBytes(k)).Encrypt(actEncr);
            byte[] expEncr = HexRead.ReadFromString(expEncrHex);

            if (!Arrays.Equals(expEncr, actEncr))
            {
                throw new ComparisonFailure("Data mismatch", HexDump.ToHex(expEncr), HexDump.ToHex(actEncr));
            }


            //Cipher cipher;
            //try
            //{
            //    cipher = Cipher.GetInstance("RC4");
            //}
            //catch (GeneralSecurityException e)
            //{
            //    throw new Exception(e);
            //}
            //String k2 = k + k; // Sun has minimum of 5 bytes for key
            //SecretKeySpec skeySpec = new SecretKeySpec(k2.Bytes, "RC4");

            //try
            //{
            //    cipher.Init(Cipher.DECRYPT_MODE, skeySpec);
            //}
            //catch (InvalidKeyException e)
            //{
            //    throw new Exception(e);
            //}
            //byte[] origData = origText.getBytes();
            //byte[] altEncr = cipher.update(origData);
            //if (!Arrays.Equals(expEncr, altEncr))
            //{
            //    throw new Exception("Mismatch from jdk provider");
            //}
        }
        
    }

}