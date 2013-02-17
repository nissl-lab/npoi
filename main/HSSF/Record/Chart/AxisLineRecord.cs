
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


namespace NPOI.HSSF.Record.Chart
{

    using System;
    using System.Text;
    using NPOI.Util;


    public enum AxisLineType
    {
        /// <summary>
        /// The axis (section 2.2.3.6) line itself.
        /// </summary>
        AxisLine,
        /// <summary>
        /// The major gridlines along the axis
        /// </summary>
        MajorGridLine,
        /// <summary>
        /// The minor gridlines along the axis
        /// </summary>
        MinorGridLine,
        /// <summary>
        /// The walls or floor of a 3-D chart
        /// </summary>
        WallsOrFloorOf3D,
    }
    /*
     * The axis line format record defines the axis type details.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     *///
    
    /// <summary>
    /// The AxisLine record specifies which part of the axis (section 2.2.3.6) is 
    /// specified by the LineFormat record (section 2.4.156) that follows.
    /// 
    /// Excel Binary File Format (.xls) Structure Specification 
    /// </summary>
    public class AxisLineRecord
       : StandardRecord
    {
        public const short sid = 0x1021;
        private short field_1_axisType;
        public const short AXIS_TYPE_AXIS_LINE = 0;
        public const short AXIS_TYPE_MAJOR_GRID_LINE = 1;
        public const short AXIS_TYPE_MINOR_GRID_LINE = 2;
        public const short AXIS_TYPE_WALLS_OR_FLOOR = 3;


        public AxisLineRecord()
        {

        }

        /**
         * Constructs a AxisLineFormat record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public AxisLineRecord(RecordInputStream in1)
        {
            field_1_axisType = in1.ReadShort();

        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[AXISLINEFORMAT]\n");
            buffer.Append("    .axisType             = ")
                .Append("0x").Append(HexDump.ToHex(AxisType))
                .Append(" (").Append(AxisType).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/AXISLINEFORMAT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_axisType);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return 2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            AxisLineRecord rec = new AxisLineRecord();

            rec.field_1_axisType = field_1_axisType;
            return rec;
        }




        /*
         * Get the axis type field for the AxisLineFormat record.
         *
         * @return  One of 
         *        AXIS_TYPE_AXIS_LINE
         *        AXIS_TYPE_MAJOR_GRID_LINE
         *        AXIS_TYPE_MINOR_GRID_LINE
         *        AXIS_TYPE_WALLS_OR_FLOOR
         *///
        
        /// <summary>
        /// 
        /// </summary>
        public short AxisType
        {
            get { return field_1_axisType; }
            set { this.field_1_axisType = value; }
        }

    }
}




