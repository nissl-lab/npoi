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

namespace TestCases.SS.Formula.PTG
{
    using System;
    using NPOI.HSSF.UserModel;
    using TestCases.HSSF;

    /**
     * Convenient abstract class to reduce the amount of boilerplate code needed
     * in ptg-related unit Tests.
     *
     * @author Daniel Noll (daniel at nuix dot com dot au)
     */
    public abstract class AbstractPtgTestCase
    {

        /**
         * Loads a workbook from the given filename in the Test data dir.
         *
         * @param sampleFileName the filename.
         * @return the loaded workbook.
         */
        protected static HSSFWorkbook LoadWorkbook(String sampleFileName)
        {
            return HSSFTestDataSamples.OpenSampleWorkbook(sampleFileName);
        }

        /**
         * Creates a new Workbook and Adds one sheet with the specified name
         */
        protected static HSSFWorkbook CreateWorkbookWithSheet(String sheetName)
        {
            HSSFWorkbook book = new HSSFWorkbook();
            book.CreateSheet(sheetName);
            return book;
        }
    }


}