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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.OpenXmlFormats.Dml;

namespace NPOI.XSSF.UserModel {
    public class XSSFChildGroupAnchor : XSSFAnchor
    {
        private CT_GroupTransform2D gt2d;
        public XSSFChildGroupAnchor(int x, int y, int cx, int cy)
        {
            gt2d = new CT_GroupTransform2D();
            CT_Point2D off = gt2d.AddNewOff();
            CT_PositiveSize2D ext = gt2d.AddNewExt();
            CT_Point2D chOff = gt2d.AddNewChOff();
            CT_PositiveSize2D chExt = gt2d.AddNewChExt();

            off.x = Math.Min(x, cx);
            off.y = Math.Min(y, cy);
            ext.cx = Math.Abs(cx - x);
            ext.cy = Math.Abs(cy - y);
            if (x > cx) gt2d.flipH = true;
            if (y > cy) gt2d.flipV = true;

            chOff.x = off.x;
            chOff.y = off.y;
            chExt.cx = ext.cx;
            chExt.cy = ext.cy;
        }

        public XSSFChildGroupAnchor(CT_GroupTransform2D gt2d)
        {
            this.gt2d = gt2d;
        }


        public CT_GroupTransform2D GetCTTransform2D()
        {
            return gt2d;
        }

        public override int Dx1
        {
            get
            {
                return (int)gt2d.off.x;

            }
            set 
            {
                gt2d.off.y = (value);
            }
        }

        public override int Dy1
        {
            get
            {
                return (int)gt2d.off.y;
            }
            set 
            {
                gt2d.off.y = (value);
            }
        }

        public override int Dy2
        {
            get
            {
                return (int)(Dy1 + gt2d.ext.cy);
            }
            set 
            {
                gt2d.ext.cy = (value - Dy1);
            }
        }

        public override int Dx2
        {
            get
            {
                return (int)(Dx1 + gt2d.ext.cx);
            }
            set 
            {
                gt2d.ext.cx = (value - Dx1);
            }
        }
    }
}
