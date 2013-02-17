
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
        

/*
 * FontFormatting.java
 *
 * Created on January 22, 2008, 10:05 PM
 */

namespace NPOI.HSSF.Record.CF
{
    using System;
    using System.Text;
    using NPOI.HSSF.Record;
    using NPOI.SS.UserModel;
    using NPOI.Util;



    /**
     * Border Formatting Block of the Conditional Formatting Rule Record.
     * 
     * @author Dmitriy Kumshayev
     */

    public class BorderFormatting
    {
        public BorderFormatting()
        {
            field_13_border_styles1 = (short)0;
            field_14_border_styles2 = (short)0;
        }

        /** Creates new FontFormatting */
        public BorderFormatting(RecordInputStream in1)
        {
            field_13_border_styles1 = in1.ReadInt();
            field_14_border_styles2 = in1.ReadInt();
        }

        // BORDER FORMATTING BLOCK
        // For Border Line Style codes see HSSFCellStyle.BORDER_XXXXXX
        private int field_13_border_styles1;
        private static BitField bordLeftLineStyle = BitFieldFactory.GetInstance(0x0000000F);
        private static BitField bordRightLineStyle = BitFieldFactory.GetInstance(0x000000F0);
        private static BitField bordTopLineStyle = BitFieldFactory.GetInstance(0x00000F00);
        private static BitField bordBottomLineStyle = BitFieldFactory.GetInstance(0x0000F000);
        private static BitField bordLeftLineColor = BitFieldFactory.GetInstance(0x007F0000);
        private static BitField bordRightLineColor = BitFieldFactory.GetInstance(0x3F800000);
        private static BitField bordTlBrLineOnOff = BitFieldFactory.GetInstance(0x40000000);
        private static BitField bordBlTrtLineOnOff = BitFieldFactory.GetInstance(unchecked((int)0x80000000));

        private int field_14_border_styles2;
        private static BitField bordTopLineColor = BitFieldFactory.GetInstance(0x0000007F);
        private static BitField bordBottomLineColor = BitFieldFactory.GetInstance(0x00003f80);
        private static BitField bordDiagLineColor = BitFieldFactory.GetInstance(0x001FC000);
        private static BitField bordDiagLineStyle = BitFieldFactory.GetInstance(0x01E00000);

        /// <summary>
        /// Get the type of border to use for the left border of the cell
        /// </summary>
        public BorderStyle BorderLeft {
            get {
                return (BorderStyle) bordLeftLineStyle.GetValue (field_13_border_styles1);
            }
            set {
                field_13_border_styles1 = bordLeftLineStyle.SetValue (field_13_border_styles1, (int) value);
            }
        }

          
        /// <summary>
        /// Get the type of border to use for the right border of the cell
        /// </summary>
        public BorderStyle BorderRight {
            get {
                return (BorderStyle) bordRightLineStyle.GetValue (field_13_border_styles1);
            }
            set {
                field_13_border_styles1 = bordRightLineStyle.SetValue (field_13_border_styles1, (int) value);
            }
        }

        /// <summary>
        /// Get the type of border to use for the top border of the cell
        /// </summary>
        public BorderStyle BorderTop {
            get {
                return (BorderStyle) bordTopLineStyle.GetValue (field_13_border_styles1);
            }
            set {
                field_13_border_styles1 = bordTopLineStyle.SetValue (field_13_border_styles1, (int) value);
            }
        }

         
        /// <summary>
        /// Get the type of border to use for the bottom border of the cell
        /// </summary>
        public BorderStyle BorderBottom {
            get {
                return (BorderStyle) bordBottomLineStyle.GetValue (field_13_border_styles1);
            }
            set {
                field_13_border_styles1 = bordBottomLineStyle.SetValue (field_13_border_styles1, (int) value);
            }
        }

        /// <summary>
        ///  Get the type of border to use for the diagonal border of the cell
        /// </summary>
        public BorderStyle BorderDiagonal {
            get {
                return (BorderStyle) bordDiagLineStyle.GetValue (field_14_border_styles2);
            }
            set {
                field_14_border_styles2 = bordDiagLineStyle.SetValue (field_14_border_styles2, (int) value);
            }
        }


        /// <summary>
        /// Get the color to use for the left border
        /// </summary>
        public short LeftBorderColor
        {
            get
            {
                return (short)bordLeftLineColor.GetValue(field_13_border_styles1);
            }
            set
            {
                field_13_border_styles1 = bordLeftLineColor.SetValue(field_13_border_styles1, value);
            }
        }

        /// <summary>
        /// Get the color to use for the right border
        /// </summary>
        public short RightBorderColor
        {
            get
            {
                return (short)bordRightLineColor.GetValue(field_13_border_styles1);
            }
            set
            {
                field_13_border_styles1 = bordRightLineColor.SetValue(field_13_border_styles1, value);
            }
        }

        /// <summary>
        /// Get the color to use for the top border
        /// </summary>
        public short TopBorderColor
        {
            get
            {
                return (short)bordTopLineColor.GetValue(field_14_border_styles2);
            }
            set
            {
                field_14_border_styles2 = bordTopLineColor.SetValue(field_14_border_styles2, value);
            }
        }

         
        /// <summary>
        /// Get the color to use for the bottom border
        /// </summary>
        public short BottomBorderColor
        {
            get
            {
                return (short)bordBottomLineColor.GetValue(field_14_border_styles2);
            }
            set
            {
                field_14_border_styles2 = bordBottomLineColor.SetValue(field_14_border_styles2, value);
            }
        }

         
        /// <summary>
        /// Get the color to use for the diagonal border
        /// </summary>
        public short DiagonalBorderColor
        {
            get
            {
                return (short)bordDiagLineColor.GetValue(field_14_border_styles2);
            }
            set
            {
                field_14_border_styles2 = bordDiagLineColor.SetValue(field_14_border_styles2, value);
            }
        }
        /// <summary>
        /// true if forward diagonal is on 
        /// </summary>
        public bool IsForwardDiagonalOn
        {
            get
            {
                return bordBlTrtLineOnOff.IsSet(field_13_border_styles1);
            }
            set
            {
                field_13_border_styles1 = bordBlTrtLineOnOff.SetBoolean(field_13_border_styles1, value);
            }
        }
        
        /// <summary>
        /// true if backward diagonal Is on
        /// </summary>
        public bool IsBackwardDiagonalOn
        {
            get
            {
                return bordTlBrLineOnOff.IsSet(field_13_border_styles1);
            }
            set
            {
                field_13_border_styles1 = bordTlBrLineOnOff.SetBoolean(field_13_border_styles1, value);
            }
        }


        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("    [Border Formatting]\n");
            buffer.Append("          .lftln     = ").Append(StringUtil.ToHexString((int) BorderLeft)).Append("\n");
            buffer.Append("          .rgtln     = ").Append(StringUtil.ToHexString((int) BorderRight)).Append("\n");
            buffer.Append("          .topln     = ").Append(StringUtil.ToHexString((int) BorderTop)).Append("\n");
            buffer.Append("          .btmln     = ").Append(StringUtil.ToHexString((int) BorderBottom)).Append("\n");
            buffer.Append("          .leftborder= ").Append(StringUtil.ToHexString(LeftBorderColor)).Append("\n");
            buffer.Append("          .rghtborder= ").Append(StringUtil.ToHexString(RightBorderColor)).Append("\n");
            buffer.Append("          .topborder= ").Append(StringUtil.ToHexString(TopBorderColor)).Append("\n");
            buffer.Append("          .bottomborder= ").Append(StringUtil.ToHexString(BottomBorderColor)).Append("\n");
            buffer.Append("          .fwdiag= ").Append(IsForwardDiagonalOn).Append("\n");
            buffer.Append("          .bwdiag= ").Append(IsBackwardDiagonalOn).Append("\n");
            buffer.Append("    [/Border Formatting]\n");
            return buffer.ToString();
        }

        public Object Clone()
        {
            BorderFormatting rec = new BorderFormatting();
            rec.field_13_border_styles1 = field_13_border_styles1;
            rec.field_14_border_styles2 = field_14_border_styles2;
            return rec;
        }

        public int Serialize(int offset, byte[] data)
        {
            LittleEndian.PutInt(data, offset, field_13_border_styles1);
            offset += 4;
            LittleEndian.PutInt(data, offset, field_14_border_styles2);
            offset += 4;
            return 8;
        }
        public void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteInt(field_13_border_styles1);
            out1.WriteInt(field_14_border_styles2);
        }
    }
}