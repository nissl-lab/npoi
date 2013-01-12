/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestCases.HSSF.UserModel
{

    /**
     * @author Evgeniy Berlog
     * @date 01.08.12
     */
    public class TestPatriarch
    {

        public void TestGetPatriarch()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet();
            Assert.IsNull(sh.DrawingPatriarch());

            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch();
            Assert.IsNotNull(patriarch);
            patriarch.CreateSimpleShape(new HSSFClientAnchor());
            patriarch.CreateSimpleShape(new HSSFClientAnchor());

            Assert.AreSame(patriarch, sh.DrawingPatriarch());

            EscherAggregate agg = patriarch.GetBoundAggregate();

            EscherDgRecord dg = agg.GetEscherContainer().GetChildById(EscherDgRecord.RECORD_ID);
            int lastId = dg.GetLastMSOSPID();

            Assert.AreSame(patriarch, sh.CreateDrawingPatriarch());

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0);
            patriarch = sh.CreateDrawingPatriarch();
            dg = patriarch.GetBoundAggregate().GetEscherContainer().GetChildById(EscherDgRecord.RECORD_ID);

            Assert.AreEqual(lastId, dg.GetLastMSOSPID());
        }
    }
}
