
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License Is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HWPF.UserModel
{
    using System;
    using NPOI.HWPF.Model.Types;

    public class TableCellDescriptor : TCAbstractType
    {
        public static int SIZE = 20;

        public TableCellDescriptor()
        {
            field_3_brcTop = new BorderCode();
            field_4_brcLeft = new BorderCode();
            field_5_brcBottom = new BorderCode();
            field_6_brcRight = new BorderCode();

        }

        //public Object Clone()
        //{
        //  TableCellDescriptor tc = (TableCellDescriptor)base.Clone();
        //  tc.field_3_brcTop = (BorderCode)field_3_brcTop.Clone();
        //  tc.field_4_brcLeft = (BorderCode)field_4_brcLeft.Clone();
        //  tc.field_5_brcBottom = (BorderCode)field_5_brcBottom.clone();
        //  tc.field_6_brcRight = (BorderCode)field_6_brcRight.clone();
        //  return tc;
        //}

        public static TableCellDescriptor ConvertBytesToTC(byte[] buf, int offset)
        {
            TableCellDescriptor tc = new TableCellDescriptor();
            tc.FillFields(buf, offset);
            return tc;
        }

    }
}