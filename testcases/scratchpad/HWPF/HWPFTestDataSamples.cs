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
using NPOI.HWPF;
using System.IO;
using NPOI.POIFS.FileSystem;
using System;
namespace TestCases.HWPF
{


    public class HWPFTestDataSamples
    {

        public static HWPFDocument OpenSampleFile(String sampleFileName)
        {
            Stream is1 = POIDataSamples.GetDocumentInstance().OpenResourceAsStream(sampleFileName);
            return new HWPFDocument(is1);

        }
        public static HWPFOldDocument OpenOldSampleFile(String sampleFileName)
        {
            Stream is1 = POIDataSamples.GetDocumentInstance().OpenResourceAsStream(sampleFileName);
            return new HWPFOldDocument(new POIFSFileSystem(is1));
        }
        /**
         * Writes a spreadsheet to a <tt>MemoryStream</tt> and Reads it back
         * from a <tt>MemoryStream</tt>.<p/>
         * Useful for verifying that the serialisation round trip
         */
        public static HWPFDocument WriteOutAndReadBack(HWPFDocument original)
        {
            MemoryStream baos = new MemoryStream(4096);
            original.Write(baos);
            MemoryStream bais = new MemoryStream(baos.ToArray());
            return new HWPFDocument(bais);

        }
    }
}
