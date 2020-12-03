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
    using System.Text;

    [TestFixture]
    public class TestEscherDump
    {
        static NullPrinterStream nullPS;
        [SetUp]
        public static void Init()
        {
            nullPS = new NullPrinterStream();
        }
        [Test, Ignore("Not Implemented")]
        public void TestSimple()
        {
            // simple test to at least cover some parts of the class
            //EscherDump.Main(new String[] {});

            //new EscherDump().Dump(0, new byte[] { }, nullPS);
            //new EscherDump().Dump(new byte[] { }, 0, 0, nullPS);
            //new EscherDump().DumpOld(0, new MemoryStream(new byte[] { }), nullPS);
        }


        [Test, Ignore("Not Implemented")]
        public void TestWithData()
        {
            //new EscherDump().DumpOld(8, new MemoryStream(new byte[] { 00, 00, 00, 00, 00, 00, 00, 00 }), nullPS);
        }

        [Test, Ignore("Not Implemented")]
        public void TestWithSamplefile()
        {
            //InputStream stream = HSSFTestDataSamples.openSampleFileStream(")
            //byte[] data = POIDataSamples.DDFInstance.readFile("Container.dat");
            //new EscherDump().dump(data.length, data, nullPS);
            ////new EscherDump().dumpOld(data.length, new ByteArrayInputStream(data), System.out);

            //data = new byte[2586114];
            //InputStream stream = HSSFTestDataSamples.openSampleFileStream("44593.xls");
            //try
            //{
            //    int bytes = IOUtils.readFully(stream, data);
            //    Assert.IsTrue(bytes != -1);
            //    //new EscherDump().dump(bytes, data, System.out);
            //    //new EscherDump().dumpOld(bytes, new ByteArrayInputStream(data), System.out);
            //}
            //finally
            //{
            //    stream.close();
            //}
        }

        /**
         * Implementation of an OutputStream which does nothing, used
         * to redirect stdout to avoid spamming the console with output
         */
        private class NullPrinterStream : TextWriter
        {
            public NullPrinterStream()
            {
                //super(new NullOutputStream(),true,LocaleUtil.CHARSET_1252.name());
            }

            public override Encoding Encoding
            {
                get { return Encoding.GetEncoding(1252); }
            }

            /**
            * Implementation of an OutputStream which does nothing, used
            * to redirect stdout to avoid spamming the console with output
            */
            //private class NullOutputStream : Stream
            //{
            //    public void write(byte[] b, int off, int len)
            //    {
            //    }
            //    public void write(int b)
            //    {
            //    }
            //    public void write(byte[] b) 
            //    {
            //    }
            //}
        }

    }
}
