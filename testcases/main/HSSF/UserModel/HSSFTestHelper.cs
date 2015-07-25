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
    using NPOI.HSSF.Model;
    using NPOI.HSSF.UserModel;
    using NPOI.DDF;
    using NPOI.HSSF.Record;

    /**
     * Helper class for HSSF Tests that aren't within the
     *  HSSF UserModel package, but need to do internal
     *  UserModel things.
     */
    public class HSSFTestHelper
    {

        public class MockDrawingManager : DrawingManager2
        {

            public MockDrawingManager()
                : base(null)
            {
            }

            public override int AllocateShapeId(short drawingGroupId)
            {
                return 1025; //Mock value
            }

            public override int AllocateShapeId(short drawingGroupId, EscherDgRecord dg)
            {
                return 1025;
            }

            public override EscherDgRecord CreateDgRecord()
            {
                EscherDgRecord dg = new EscherDgRecord();
                dg.RecordId = (EscherDgRecord.RECORD_ID);
                dg.Options = ((short)(16));
                dg.NumShapes = (1);
                dg.LastMSOSPID = (1024);
                return dg;
            }
        }

        /**
         * Lets non UserModel Tests at the low level Workbook
         */
        public static InternalWorkbook GetWorkbookForTest(HSSFWorkbook wb)
        {
            return wb.Workbook;
        }
        public static InternalSheet GetSheetForTest(HSSFSheet sheet)
        {
            return sheet.Sheet;
        }

        public static HSSFPatriarch CreateTestPatriarch(HSSFSheet sheet, EscherAggregate agg)
        {
            return new HSSFPatriarch(sheet, agg);
        }

        public static EscherAggregate GetEscherAggregate(HSSFPatriarch patriarch)
        {
            return patriarch.GetBoundAggregate();
        }

        public static int AllocateNewShapeId(HSSFPatriarch patriarch)
        {
            return patriarch.NewShapeId();
        }

        public static EscherOptRecord GetOptRecord(HSSFShape shape)
        {
            return shape.GetOptRecord();
        }

        public static void SetShapeId(HSSFShape shape, int id)
        {
            shape.ShapeId = (id);
        }

        public static EscherContainerRecord GetEscherContainer(HSSFShape shape)
        {
            return shape.GetEscherContainer();
        }

        public static TextObjectRecord GetTextObjRecord(HSSFSimpleShape shape)
        {
            return shape.GetTextObjectRecord();
        }

        public static ObjRecord GetObjRecord(HSSFShape shape)
        {
            return shape.GetObjRecord();
        }

        public static EscherRecord GetEscherAnchor(HSSFAnchor anchor)
        {
            return anchor.GetEscherAnchor();
        }
    }

}