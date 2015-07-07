/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF
{
    using System;
    using System.IO;
    using System.Text;
    using NPOI.HSSF.UserModel;

    /**
     * Centralises logic for finding/opening sample files in the src/testcases/org/apache/poi/hssf/hssf/data folder. 
     * 
     * @author Josh Micich
     */
    public class HSSFTestDataSamples
    {

        private static POIDataSamples _inst = POIDataSamples.GetSpreadSheetInstance();

        public static FileInfo GetSampleFile(string sampleFileName)
        {
            return _inst.GetFileInfo(sampleFileName);
        }
        public static Stream OpenSampleFileStream(String sampleFileName)
        {
            return _inst.OpenResourceAsStream(sampleFileName);
        }
        public static byte[] GetTestDataFileContent(String fileName)
        {
            return _inst.ReadFile(fileName);
        }
        public static HSSFWorkbook OpenSampleWorkbook(String sampleFileName)
        {
            using (var sampleStream = _inst.OpenResourceAsStream(sampleFileName))
            {
                return new HSSFWorkbook(sampleStream);
            }
        }
        /**
         * Writes a spReadsheet to a <c>MemoryStream</c> and Reads it back
         * from a <c>ByteArrayStream</c>.<p/>
         * Useful for verifying that the serialisation round trip
         */
        public static HSSFWorkbook WriteOutAndReadBack(HSSFWorkbook original)
        {
            using (MemoryStream baos = new MemoryStream(4096))
            {
                original.Write(baos);
                return new HSSFWorkbook(baos);
            }
        }


    }
}