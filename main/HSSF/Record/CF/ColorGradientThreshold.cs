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
    using NPOI.Util;

    /**
     * Color Gradient / Color Scale specific Threshold / value (CFVO),
     *  for Changes in Conditional Formatting
     */
    public class ColorGradientThreshold : Threshold, ICloneable
    {
        private double position;

        public ColorGradientThreshold() : base()
        {

            position = 0d;
        }

        /** Creates new Color Gradient Threshold */
        public ColorGradientThreshold(ILittleEndianInput in1) : base(in1)
        {

            position = in1.ReadDouble();
        }

        public double Position
        {
            get { return position; }
            set { this.position = value; }
        }

        public override int DataLength
        {
            get { return base.DataLength + 8; }
        }

        public Object Clone()
        {
            ColorGradientThreshold rec = new ColorGradientThreshold();
            base.CopyTo(rec);
            rec.position = position;
            return rec;
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            base.Serialize(out1);
            out1.WriteDouble(position);
        }
    }

}