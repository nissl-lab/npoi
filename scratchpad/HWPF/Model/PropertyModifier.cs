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
using NPOI.Util;
using System;
using System.Text;
namespace NPOI.HWPF.Model
{


    public class PropertyModifier
    {
        /**
         * <li>"Set to 0 for variant 1" <li>"Set to 1 for variant 2"
         */
        private static BitField _fComplex = new BitField(0x0001);

        /**
         * "Index to a grpprl stored in CLX portion of file"
         */
        private static BitField _figrpprl = new BitField(0xfffe);

        /**
         * "Index to entry into rgsprmPrm"
         */
        private static BitField _fisprm = new BitField(0x00fe);

        /**
         * "sprm's operand"
         */
        private static BitField _fval = new BitField(0xff00);

        private short value;

        public PropertyModifier(short value)
        {
            this.value = value;
        }


        protected PropertyModifier Clone()
        {
            return new PropertyModifier(value);
        }


        public override bool Equals(Object obj)
        {
            if (this == obj)
                return true;
            if (obj == null)
                return false;
            if (this.GetType() != obj.GetType())
                return false;
            PropertyModifier other = (PropertyModifier)obj;
            if (value != other.value)
                return false;
            return true;
        }

        /**
         * "Index to a grpprl stored in CLX portion of file"
         */
        public short GetIgrpprl()
        {
            if (!IsComplex())
                throw new InvalidOperationException("Not complex");

            return _figrpprl.GetShortValue(value);
        }

        public short GetIsprm()
        {
            if (IsComplex())
                throw new InvalidOperationException("Not simple");

            return _fisprm.GetShortValue(value);
        }

        public short GetVal()
        {
            if (IsComplex())
                throw new InvalidOperationException("Not simple");

            return _fval.GetShortValue(value);
        }

        public short GetValue()
        {
            return value;
        }


        public override int GetHashCode()
        {
            int prime = 31;
            int result = 1;
            result = prime * result + value;
            return result;
        }

        public bool IsComplex()
        {
            return _fComplex.IsSet(value);
        }


        public override String ToString()
        {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.Append("[PRM] (complex: ");
            stringBuilder.Append(IsComplex());
            stringBuilder.Append("; ");
            if (IsComplex())
            {
                stringBuilder.Append("igrpprl: ");
                stringBuilder.Append(GetIgrpprl());
                stringBuilder.Append("; ");
            }
            else
            {
                stringBuilder.Append("isprm: ");
                stringBuilder.Append(GetIsprm());
                stringBuilder.Append("; ");
                stringBuilder.Append("val: ");
                stringBuilder.Append(GetVal());
                stringBuilder.Append("; ");
            }
            stringBuilder.Append(")");
            return stringBuilder.ToString();
        }
    }


}