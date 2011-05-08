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

namespace TestCases.HSSF.UserModel
{
    using System;
    using System.IO;
    using NPOI.HSSF.UserModel;

    using TestCases.HSSF;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.HSSF.Record;

    /**
     * @author aviks
     * 
     * This Testcase contains Tests for bugs that are yet to be fixed. Therefore,
     * the standard ant Test target does not run these Tests. Run this Testcase with
     * the single-Test target. The names of the Tests usually correspond to the
     * Bugzilla id's PLEASE MOVE Tests from this class to TestBugs once the bugs are
     * fixed, so that they are then run automatically.
     */
    [TestClass]
    public class TestUnfixedBugs
    {
        //In POI bugzilla, this bug is taged as "RESOLVED WON'T FIX"
        [TestMethod]
        public void Test43493()
        {
            // Has crazy corrupt sub-records on
            // a EmbeddedObjectRefSubRecord
            try
            {
                HSSFTestDataSamples.OpenSampleWorkbook("43493.xls");
            }
            catch (RecordFormatException e)
            {
                if (e.InnerException.InnerException is IndexOutOfRangeException)
                {
                    throw new AssertFailedException("Identified bug 43493");
                }
                throw e;
            }
        }
    }
}