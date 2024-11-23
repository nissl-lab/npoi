/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for Additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

namespace NPOI.XSSF
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using NPOI.SS;
    using NPOI.SS.UserModel;
    using NPOI.XSSF.Streaming;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using TestCases;
    using TestCases.SS;

    /**
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class SXSSFITestDataProvider : ITestDataProvider
    {
        public static SXSSFITestDataProvider instance = new SXSSFITestDataProvider();

        // an instance of all SXSSFWorkbooks opened by this TestDataProvider,
        // so that the temporary files created can be disposed up by cleanup() 
        private List<SXSSFWorkbook> instances = new List<SXSSFWorkbook>();

        private SXSSFITestDataProvider()
        {
            // enforce Singleton
        }

        public IWorkbook OpenSampleWorkbook(String sampleFileName)
        {
            XSSFWorkbook xssfWorkbook = XSSFITestDataProvider.instance.OpenSampleWorkbook(sampleFileName) as XSSFWorkbook;
            SXSSFWorkbook swb = new SXSSFWorkbook(xssfWorkbook);
            instances.Add(swb);
            return swb;
        }

        public IWorkbook WriteOutAndReadBack(IWorkbook wb)
        {
            // wb is usually an SXSSFWorkbook, but must also work on an XSSFWorkbook
            // since workbooks must be able to be written out and read back
            // several times in succession
            if (!(wb is SXSSFWorkbook || wb is XSSFWorkbook))
            {
                throw new ArgumentException("Expected an instance of XSSFWorkbook or SXSSFWorkbook");
            }

            XSSFWorkbook result;
            try
            {
                MemoryStream baos = new MemoryStream(8192);
                wb.Write(baos, false);
                Stream is1 = new MemoryStream(baos.ToArray());
                result = new XSSFWorkbook(is1);
            }
            catch (IOException e)
            {
                throw new Exception(e.Message, e);
            }
            return result;
        }

        public IWorkbook CreateWorkbook()
        {
            SXSSFWorkbook wb = new SXSSFWorkbook();
            instances.Add(wb);
            return wb;
        }

        //************ SXSSF-specific methods ***************//
        public IWorkbook CreateWorkbook(int rowAccessWindowSize)
        {
            SXSSFWorkbook wb = new SXSSFWorkbook(rowAccessWindowSize);
            instances.Add(wb);
            return wb;
        }
        public void TrackAllColumnsForAutosizing(ISheet sheet)
        {
            ((SXSSFSheet)sheet).TrackAllColumnsForAutoSizing();
        }
        //************ End SXSSF-specific methods ***************//
        public IFormulaEvaluator CreateFormulaEvaluator(IWorkbook wb)
        {
            return new XSSFFormulaEvaluator(((SXSSFWorkbook)wb).XssfWorkbook);
        }

        public byte[] GetTestDataFileContent(String fileName)
        {
            return POIDataSamples.GetSpreadSheetInstance().ReadFile(fileName);
        }

        public SpreadsheetVersion GetSpreadsheetVersion()
        {
            return SpreadsheetVersion.EXCEL2007;
        }

        public String StandardFileNameExtension
        {
            get
            {
                return "xlsx";
            }
        }

        public bool Cleanup()
        {
            bool ok = true;
            for (int i = 0; i < instances.Count; i++)
            {
                SXSSFWorkbook wb = instances[(i)];
                ok = ok && wb.Dispose();
            }

            instances.Clear();
            return ok;
        }
    }

}
