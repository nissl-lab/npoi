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
    using NUnit.Framework;
    using System.IO;


    [TestFixture]
    public class TestEscherDump
    {

        [Test]
        public void TestSimple()
        {
            // simple test to at least cover some parts of the class
            //EscherDump.Main(new String[] {});

            //new EscherDump().Dump(0, new byte[] { }, System.Console.Out);
            //new EscherDump().Dump(new byte[] { }, 0, 0, System.Console.Out);
            //new EscherDump().DumpOld(0, new MemoryStream(new byte[] { }), System.Console.Out);
        }


        [Test]
        public void TestWithData()
        {
            //new EscherDump().DumpOld(8, new MemoryStream(new byte[] { 00, 00, 00, 00, 00, 00, 00, 00 }), System.Console.Out);
        }
    }
}
