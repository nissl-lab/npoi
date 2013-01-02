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

using NPOI.OpenXml4Net.OPC;
using System;
using NPOI.XSSF.UserModel;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.Util;
using NPOI.HSSF;
using TestCases.HSSF;
namespace NPOI.XSSF
{

    /**
     * Centralises logic for Finding/opening sample files in the src/testcases/org/apache/poi/hssf/hssf/data folder. 
     * 
     * @author Josh Micich
     */
    public class XSSFTestDataSamples
    {

        public static OPCPackage OpenSamplePackage(String sampleName)
        {
            return OPCPackage.Open(
                  HSSFTestDataSamples.OpenSampleFileStream(sampleName)
            );
        }
        public static XSSFWorkbook OpenSampleWorkbook(String sampleName)
        {
            Stream is1 = HSSFTestDataSamples.OpenSampleFileStream(sampleName);
            return new XSSFWorkbook(is1);

        }
        public static IWorkbook WriteOutAndReadBack(IWorkbook wb)
        {
            IWorkbook result;
            try
            {
                using (MemoryStream baos = new MemoryStream(8192))
                {
                    wb.Write(baos);
                    using (Stream is1 = new MemoryStream(baos.ToArray()))
                    {
                        if (wb is HSSFWorkbook)
                        {
                            result = new HSSFWorkbook(is1);
                        }
                        else if (wb is XSSFWorkbook)
                        {
                            result = new XSSFWorkbook(is1);
                        }
                        else
                        {
                            throw new RuntimeException("Unexpected workbook type ("
                                    + wb.GetType().Name + ")");
                        }
                    }
                }
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }
            return result;
        }
    }

}

