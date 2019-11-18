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
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using System;
    using System.Text;


    /**
     * Icon / Multi-State Conditional Formatting Rule Record.
     */
    public class IconMultiStateFormatting : ICloneable
    {
        //private static POILogger log = POILogFactory.GetLogger(typeof(IconMultiStateFormatting));

        private IconSet iconSet;
        private byte options;
        private Threshold[] thresholds;

        private static BitField iconOnly = BitFieldFactory.GetInstance(0x01);
        private static BitField reversed = BitFieldFactory.GetInstance(0x04);

        public IconMultiStateFormatting()
        {
            iconSet = IconSet.GYR_3_TRAFFIC_LIGHTS;
            options = 0;
            thresholds = new Threshold[iconSet.num];
        }
        public IconMultiStateFormatting(ILittleEndianInput in1)
        {
            in1.ReadShort(); // Ignored
            in1.ReadByte();  // Reserved
            int num = in1.ReadByte();
            int set = in1.ReadByte();
            iconSet = IconSet.ById(set);
            if (iconSet.num != num)
            {
                //log.Log(POILogger.WARN, "Inconsistent Icon Set defintion, found " + iconSet + " but defined as " + num + " entries");
            }
            options = (byte)in1.ReadByte();

            thresholds = new Threshold[iconSet.num];
            for (int i = 0; i < thresholds.Length; i++)
            {
                thresholds[i] = new IconMultiStateThreshold(in1);
            }
        }

        public IconSet IconSet
        {
            get { return iconSet; }
            set { this.iconSet = value; }
        }

        public Threshold[] Thresholds
        {
            get { return thresholds; }
            set { this.thresholds = (value == null) ? null : (Threshold[])value.Clone(); }
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

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("    [Icon Formatting]\n");
            buffer.Append("          .icon_set = ").Append(iconSet).Append("\n");
            buffer.Append("          .icon_only= ").Append(IsIconOnly).Append("\n");
            buffer.Append("          .reversed = ").Append(IsReversed).Append("\n");
            foreach (Threshold t in thresholds)
            {
                buffer.Append(t.ToString());
            }
            buffer.Append("    [/Icon Formatting]\n");
            return buffer.ToString();
        }

        public Object Clone()
        {
            IconMultiStateFormatting rec = new IconMultiStateFormatting();
            rec.iconSet = iconSet;
            rec.options = options;
            rec.thresholds = new Threshold[thresholds.Length];
            Array.Copy(thresholds, 0, rec.thresholds, 0, thresholds.Length);
            return rec;
        }

        public int DataLength
        {
            get
            {
                int len = 6;
                foreach (Threshold t in thresholds)
                {
                    len += t.DataLength;
                }
                return len;
            }
        }

        public void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(0);
            out1.WriteByte(0);
            out1.WriteByte(iconSet.num);
            out1.WriteByte(iconSet.id);
            out1.WriteByte(options);
            foreach (Threshold t in thresholds)
            {
                t.Serialize(out1);
            }
        }
    }

}