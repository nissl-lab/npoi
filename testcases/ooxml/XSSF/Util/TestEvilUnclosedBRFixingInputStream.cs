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

using System.IO;
using NUnit.Framework;
using System.Text;
using NPOI.Util;
using NPOI.XSSF.Util;
using System;

namespace TestCases.XSSF.Util
{
    [TestFixture]
    [Obsolete]
    public class TestEvilUnclosedBRFixingInputStream
    {
        [Test]
        public void TestOK()
        {
            byte[] ok = Encoding.UTF8.GetBytes("<p><div>Hello There!</div> <div>Tags!</div></p>");

            EvilUnclosedBRFixingInputStream inp = new EvilUnclosedBRFixingInputStream(
                  new ByteArrayInputStream(ok)
            );

            Assert.IsTrue(Arrays.Equals(ok, IOUtils.ToByteArray(inp)));
            inp.Close();
        }
        [Test]
        public void TestProblem()
        {
            byte[] orig = Encoding.UTF8.GetBytes("<p><div>Hello<br>There!</div> <div>Tags!</div></p>");
            byte[] fixed1 = Encoding.UTF8.GetBytes("<p><div>Hello<br/>There!</div> <div>Tags!</div></p>");

            EvilUnclosedBRFixingInputStream inp = new EvilUnclosedBRFixingInputStream(
                  new ByteArrayInputStream(orig)
            );

            Assert.IsTrue(Arrays.Equals(fixed1, IOUtils.ToByteArray(inp)));
            inp.Close();
        }

        /**
         * Checks that we can copy with br tags around the buffer boundaries
         */
        [Test]
        public void TestBufferSize()
        {
            byte[] orig = Encoding.UTF8.GetBytes("<p><div>Hello<br> <br>There!</div> <div>Tags!<br><br></div></p>");
            byte[] fixed1 = Encoding.UTF8.GetBytes("<p><div>Hello<br/> <br/>There!</div> <div>Tags!<br/><br/></div></p>");

            // Vary the buffer size, so that we can end up with the br in the
            //  overflow or only part in the buffer
            for (int i = 5; i < orig.Length; i++)
            {
                EvilUnclosedBRFixingInputStream inp = new EvilUnclosedBRFixingInputStream(
                      new ByteArrayInputStream(orig)
                );

                MemoryStream bout = new MemoryStream();
                bool going = true;
                while (going)
                {
                    byte[] b = new byte[i];
                    int r = inp.Read(b);
                    if (r > 0)
                    {
                        bout.Write(b, 0, r);
                    }
                    else
                    {
                        going = false;
                    }
                }

                byte[] result = bout.ToArray();
                Assert.IsTrue(Arrays.Equals(fixed1, result));
                inp.Close();
            }
        }
    }
}

