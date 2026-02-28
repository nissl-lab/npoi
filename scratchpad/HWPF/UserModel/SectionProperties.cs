
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
    using System.Reflection;
    using NPOI.HWPF.Model.Types;

    public class SectionProperties
      : SEPAbstractType
    {
        public SectionProperties()
        {
            field_20_brcTop = new BorderCode();
            field_21_brcLeft = new BorderCode();
            field_22_brcBottom = new BorderCode();
            field_23_brcRight = new BorderCode();
            field_26_dttmPropRMark = new DateAndTime();
        }

        //public Object Clone()
        //{
        //  SectionProperties copy = (SectionProperties)base.Clone();
        //  copy.field_20_brcTop = (BorderCode)field_20_brcTop.Clone();
        //  copy.field_21_brcLeft = (BorderCode)field_21_brcLeft.Clone();
        //  copy.field_22_brcBottom = (BorderCode)field_22_brcBottom.clone();
        //  copy.field_23_brcRight = (BorderCode)field_23_brcRight.clone();
        //  copy.field_26_dttmPropRMark = (DateAndTime)field_26_dttmPropRMark.clone();

        //  return copy;
        //}

        public override bool Equals(Object obj)
        {
            FieldInfo[] fields = typeof(SectionProperties).BaseType.GetFields();
            try
            {
                for (int x = 0; x < fields.Length; x++)
                {
                    Object obj1 = fields[x].GetValue(this);
                    Object obj2 = fields[x].GetValue(obj);
                    if (obj1 == null && obj2 == null)
                    {
                        continue;
                    }
                    if (!obj1.Equals(obj2))
                    {
                        return false;
                    }
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

    }
}