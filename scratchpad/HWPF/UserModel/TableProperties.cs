
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


    public class TableProperties : TAPAbstractType
    //  implements Cloneable
    {

        public TableProperties()
        {

        }
        public TableProperties(int columns)
        {
            field_7_itcMac = (short)columns;
            field_10_rgshd = new ShadingDescriptor[columns];
            for (int x = 0; x < columns; x++)
            {
                field_10_rgshd[x] = new ShadingDescriptor();
            }
            field_11_brcBottom = new BorderCode();
            field_12_brcTop = new BorderCode();
            field_13_brcLeft = new BorderCode();
            field_14_brcRight = new BorderCode();
            field_15_brcVertical = new BorderCode();
            field_16_brcHorizontal = new BorderCode();
            field_8_rgdxaCenter = new short[columns];
            field_9_rgtc = new TableCellDescriptor[columns];
            for (int x = 0; x < columns; x++)
            {
                field_9_rgtc[x] = new TableCellDescriptor();
            }
        }

        //public Object Clone()
        //{
        //  TableProperties tap = (TableProperties)super.clone();
        //  tap.field_10_rgshd = new ShadingDescriptor[field_10_rgshd.Length];
        //  for (int x = 0; x < field_10_rgshd.Length; x++)
        //  {
        //    tap.field_10_rgshd[x] = (ShadingDescriptor)field_10_rgshd[x].clone();
        //  }
        //  tap.field_11_brcBottom = (BorderCode)field_11_brcBottom.clone();
        //  tap.field_12_brcTop = (BorderCode)field_12_brcTop.clone();
        //  tap.field_13_brcLeft = (BorderCode)field_13_brcLeft.clone();
        //  tap.field_14_brcRight = (BorderCode)field_14_brcRight.clone();
        //  tap.field_15_brcVertical = (BorderCode)field_15_brcVertical.clone();
        //  tap.field_16_brcHorizontal = (BorderCode)field_16_brcHorizontal.clone();
        //  tap.field_8_rgdxaCenter = (short[])field_8_rgdxaCenter.clone();
        //  tap.field_9_rgtc = new TableCellDescriptor[field_9_rgtc.Length];
        //  for (int x = 0; x < field_9_rgtc.Length; x++)
        //  {
        //    tap.field_9_rgtc[x] = (TableCellDescriptor)field_9_rgtc[x].clone();
        //  }
        //  return tap;
        //}

    }
}