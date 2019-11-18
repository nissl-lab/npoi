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
     * Color Gradient / Color Scale Conditional Formatting Rule Record.
     * (Called Color Gradient in the file format docs, but more commonly
     *  Color Scale in the UI)
     */
    public class ColorGradientFormatting : ICloneable
    {
        //private static POILogger log = POILogFactory.GetLogger(typeof(ColorGradientFormatting));

        private byte options = 0;
        private ColorGradientThreshold[] thresholds;
        private ExtendedColor[] colors;

        private static BitField clamp = BitFieldFactory.GetInstance(0x01);
        private static BitField background = BitFieldFactory.GetInstance(0x02);

        public ColorGradientFormatting()
        {
            options = 3;
            thresholds = new ColorGradientThreshold[3];
            colors = new ExtendedColor[3];
        }
        public ColorGradientFormatting(ILittleEndianInput in1)
        {
            in1.ReadShort(); // Ignored
            in1.ReadByte();  // Reserved
            int numI = in1.ReadByte();
            int numG = in1.ReadByte();
            if (numI != numG)
            {
                //log.Log(POILogger.WARN, "Inconsistent Color Gradient defintion, found " + numI + " vs " + numG + " entries");
            }
            options = (byte)in1.ReadByte();

            thresholds = new ColorGradientThreshold[numI];
            for (int i = 0; i < thresholds.Length; i++)
            {
                thresholds[i] = new ColorGradientThreshold(in1);
            }
            colors = new ExtendedColor[numG];
            for (int i = 0; i < colors.Length; i++)
            {
                in1.ReadDouble(); // Slightly pointless step counter
                colors[i] = new ExtendedColor(in1);
            }
        }

        public int NumControlPoints
        {
            get { return thresholds.Length; }
            set
            {
                if (value != thresholds.Length)
                {
                    ColorGradientThreshold[] nt = new ColorGradientThreshold[value];
                    ExtendedColor[] nc = new ExtendedColor[value];

                    int copy = Math.Min(thresholds.Length, value);
                    Array.Copy(thresholds, 0, nt, 0, copy);
                    Array.Copy(colors, 0, nc, 0, copy);

                    this.thresholds = nt;
                    this.colors = nc;

                    updateThresholdPositions();
                }
            }
        }

        public ColorGradientThreshold[] Thresholds
        {
            get { return thresholds; }
            set {
                this.thresholds = value == null ? null : (ColorGradientThreshold[])value.Clone(); ;
                updateThresholdPositions();
            }
        }

        public ExtendedColor[] Colors
        {
            get { return colors; }
            set { this.colors = (value == null) ? null : (ExtendedColor[])value.Clone(); }
            
        }

        public bool IsClampToCurve
        {
            get { return GetOptionFlag(clamp); }
        }
        public bool IsAppliesToBackground
        {
            get { return GetOptionFlag(background); }
        }
        private bool GetOptionFlag(BitField field)
        {
            int value = field.GetValue(options);
            return value == 0 ? false : true;
        }

        private void updateThresholdPositions()
        {
            double step = 1d / (thresholds.Length - 1);
            for (int i = 0; i < thresholds.Length; i++)
            {
                thresholds[i].Position = (/*setter*/step * i);
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("    [Color Gradient Formatting]\n");
            buffer.Append("          .clamp     = ").Append(IsClampToCurve).Append("\n");
            buffer.Append("          .background= ").Append(IsAppliesToBackground).Append("\n");
            foreach (Threshold t in thresholds)
            {
                buffer.Append(t.ToString());
            }
            foreach (ExtendedColor c in colors)
            {
                buffer.Append(c.ToString());
            }
            buffer.Append("    [/Color Gradient Formatting]\n");
            return buffer.ToString();
        }

        public Object Clone()
        {
            ColorGradientFormatting rec = new ColorGradientFormatting();
            rec.options = options;
            rec.thresholds = new ColorGradientThreshold[thresholds.Length];
            rec.colors = new ExtendedColor[colors.Length];
            Array.Copy(thresholds, 0, rec.thresholds, 0, thresholds.Length);
            Array.Copy(colors, 0, rec.colors, 0, colors.Length);
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
                foreach (ExtendedColor c in colors)
                {
                    len += c.DataLength;
                    len += 8;
                }
                return len;
            }
            
        }

        public void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(0);
            out1.WriteByte(0);
            out1.WriteByte(thresholds.Length);
            out1.WriteByte(thresholds.Length);
            out1.WriteByte(options);

            foreach (ColorGradientThreshold t in thresholds)
            {
                t.Serialize(out1);
            }

            double step = 1d / (colors.Length - 1);
            for (int i = 0; i < colors.Length; i++)
            {
                out1.WriteDouble(i * step);

                ExtendedColor c = colors[i];
                c.Serialize(out1);
            }
        }
    }

}