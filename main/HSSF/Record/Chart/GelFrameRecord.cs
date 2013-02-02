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

//using NPOI.HSSF.Record.Drawing;
//using NPOI.SS.UserModel.Drawing;

namespace NPOI.HSSF.Record.Chart
{
    /// <summary>
    /// specifies the properties of a fill pattern for parts of a chart.
    /// </summary>
    /// <remarks>
    /// author: Antony liu (antony.apollo at gmail.com)
    /// </remarks>
    public class GelFrameRecord : RowDataRecord
    {
        

        //private OfficeArtFOPT fillOption;
        //private OfficeArtTertiaryFOPT tertiaryFillOption;
        public const short sid = 0x1066;

        public GelFrameRecord(RecordInputStream ris)
            :base(ris)
        {
            //fillOption = new OfficeArtFOPT(ris);
            //tertiaryFillOption = new OfficeArtTertiaryFOPT(ris);
        }

        protected override int DataSize
        {
            get 
            {
                return base.DataSize;
                //return fillOption.DataSize + tertiaryFillOption.DataSize; 
            }
        }

        public override void Serialize(NPOI.Util.ILittleEndianOutput out1)
        {
            base.Serialize(out1);
            //fillOption.Serialize(out1);
            //tertiaryFillOption.Serialize(out1);
        }

        public override short Sid
        {
            get { return sid; }
        }
/*
        public MSOFillType FillType
        {
            get;
            set;
        }
        
        public int FillColor
        {
            get;
            set;
        }
        public int FillOpacity
        {
            ///Value of the real number = Integral + (Fractional / 65536.0)
            get;
            set;
        }
        public int FillBackColor
        {
            get;
            set;
        }

        public int FillBackOpacity
        {
            get;
            set;
        }
        public int FillCrMod
        {
            get;
            set;
        }
        public int FillBlip_complex
        {
            get;
            set;
        }
        public int FillBlipName_complex
        {
            get;
            set;
        }
        public int FillBlipFlags
        {
            get;
            set;
        }
        public int FillWidth
        {
            get;
            set;
        }
        public int FillHeight
        {
            get;
            set;
        }
        
        public int FillAngle
        {
            get;
            set;
        }
        
        public int FillFocus
        {
            get;
            set;
        }

        public int FillToLeft
        {
            get;
            set;
        }
        public int FillToTop
        {
            get;
            set;
        }
        public int FillToRight
        {
            get;
            set;
        }
        public int FillToBottom
        {
            get;
            set;
        }
        public int FillRectLeft
        {
            get;
            set;
        }
        public int FillRectTop
        {
            get;
            set;
        }
        public int FillRectRight
        {
            get;
            set;
        }
        public int FillRectBottom
        {
            get;
            set;
        }
        public int FillDztype
        {
            get;
            set;
        }
        public int FillShadePreset
        {
            get;
            set;
        }
        public int FillShadeColors_complex
        {
            get;
            set;
        }
        public int FillOriginX
        {
            get;
            set;
        }
        public int FillOriginY
        {
            get;
            set;
        }
        
        public int FillShapeOriginX
        {
            get;
            set;
        }
        public int FillShapeOriginY
        {
            get;
            set;
        }
        
        public int FillShadeType
        {
            get;
            set;
        }


        public int fFilled
        {
            get;
            set;
        }
        public int fHitTestFill
        {
            get;
            set;
        }
        public int FillShape
        {
            get;
            set;
        }
        public int FillUseRect
        {
            get;
            set;
        }
        public int fNoFillHitTest
        {
            get;
            set;
        }


        public int FillColorExt
        {
            get;
            set;
        }
        public int FillColorExtMod
        {
            get;
            set;
        }
        public int FillBackColorExt
        {
            get;
            set;
        }
        
        public int FillBackColorExtMod
        {
            get;
            set;
        }


        public int fRecolorFillAsPicture
        {
            get;
            set;
        }
        
        public int fUseShapeAnchor
        {
            get;
            set;
        }
        */
    }
}
