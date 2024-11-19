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

namespace NPOI.HSSF.Record.CF
{
    using System;
    using System.Text;
    using NPOI.HSSF.Record.Common;
    using NPOI.Util;


    /**
     * Data Bar Conditional Formatting Rule Record.
     */
    public class DataBarFormatting : ICloneable
    {

        private byte options = 0;
        private byte percentMin = 0;
        private byte percentMax = 0;
        private ExtendedColor color;
        private DataBarThreshold thresholdMin;
        private DataBarThreshold thresholdMax;

        private static BitField iconOnly = BitFieldFactory.GetInstance(0x01);
        private static BitField reversed = BitFieldFactory.GetInstance(0x04);

        public DataBarFormatting()
        {
            options = 2;
        }
        public DataBarFormatting(ILittleEndianInput in1)
        {
            in1.ReadShort(); // Ignored
            in1.ReadByte();  // Reserved
            options = (byte)in1.ReadByte();

            percentMin = (byte)in1.ReadByte();
            percentMax = (byte)in1.ReadByte();
            if (percentMin < 0 || percentMin > 100)
                //log.Log(POILogger.WARN, "Inconsistent Minimum Percentage found " + percentMin);
                Console.WriteLine("Inconsistent Minimum Percentage found " + percentMin);

            if (percentMax < 0 || percentMax > 100)
                //log.Log(POILogger.WARN, "Inconsistent Minimum Percentage found " + percentMin);
                Console.WriteLine("Inconsistent Maximum Percentage found " + percentMax);

            color = new ExtendedColor(in1);
            thresholdMin = new DataBarThreshold(in1);
            thresholdMax = new DataBarThreshold(in1);
        }

        public bool IsIconOnly
        {
            get { return GetOptionFlag(iconOnly); }
            set { SetOptionFlag(value, iconOnly); }
        }

        public bool IsReversed
        {
            get { return GetOptionFlag(reversed); }
            set { SetOptionFlag(value, reversed); }
        }

        private bool GetOptionFlag(BitField field)
        {
            int value = field.GetValue(options);
            return value == 0 ? false : true;
        }
        private void SetOptionFlag(bool option, BitField field)
        {
            options = field.SetByteBoolean(options, option);
        }

        public byte PercentMin
        {
            get { return percentMin; }
            set { this.percentMin = value; }
        }

        public byte PercentMax
        {
            get { return percentMax; }
            set { this.percentMax = value; }
        }

        public ExtendedColor Color
        {
            get { return color; }
            set { this.color = value; }
        }

        public DataBarThreshold ThresholdMin
        {
            get { return thresholdMin; }
            set { this.thresholdMin = value; }
        }


        public DataBarThreshold ThresholdMax
        {
            get { return thresholdMax; }
            set { this.thresholdMax = value; }
        }


        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("    [Data Bar Formatting]\n");
            buffer.Append("          .icon_only= ").Append(IsIconOnly).Append("\n");
            buffer.Append("          .reversed = ").Append(IsReversed).Append("\n");
            buffer.Append(color);
            buffer.Append(thresholdMin);
            buffer.Append(thresholdMax);
            buffer.Append("    [/Data Bar Formatting]\n");
            return buffer.ToString();
        }

        public Object Clone()
        {
            DataBarFormatting rec = new DataBarFormatting();
            rec.options = options;
            rec.percentMin = percentMin;
            rec.percentMax = percentMax;
            rec.color = (ExtendedColor)color.Clone();
            rec.thresholdMin = (DataBarThreshold)thresholdMin.Clone();
            rec.thresholdMax = (DataBarThreshold)thresholdMax.Clone();
            return rec;
        }

        public int DataLength
        {
            get
            {
                return 6 + color.DataLength +
                   thresholdMin.DataLength +
                   thresholdMax.DataLength;
            }
            
        }

        public void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(0);
            out1.WriteByte(0);
            out1.WriteByte(options);
            out1.WriteByte(percentMin);
            out1.WriteByte(percentMax);
            color.Serialize(out1);
            thresholdMin.Serialize(out1);
            thresholdMax.Serialize(out1);
        }
    }

}