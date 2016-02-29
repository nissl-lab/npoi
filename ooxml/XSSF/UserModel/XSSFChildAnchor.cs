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

using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using System;
using NPOI.OpenXmlFormats.Dml;
namespace NPOI.XSSF.UserModel
{
    public class XSSFChildAnchor : XSSFAnchor
    {
        private CT_Transform2D t2d;

        public XSSFChildAnchor(int x, int y, int cx, int cy)
        {
            t2d = new CT_Transform2D();
            CT_Point2D off = t2d.AddNewOff();
            CT_PositiveSize2D ext = t2d.AddNewExt();

            off.x = (x);
            off.y = (y);
            ext.cx = (Math.Abs(cx - x));
            ext.cy = (Math.Abs(cy - y));
            if (x > cx) t2d.flipH = (true);
            if (y > cy) t2d.flipV = (true);
        }

        public XSSFChildAnchor(CT_Transform2D t2d)
        {
            this.t2d = t2d;
        }


        public CT_Transform2D GetCTTransform2D()
        {
            return t2d;
        }

        public override int Dx1
        {
            get
            {
                return (int)t2d.off.x;

            }
            set 
            {
                t2d.off.y = (value);
            }
        }

        public override int Dy1
        {
            get
            {
                return (int)t2d.off.y;
            }
            set 
            {
                t2d.off.y = (value);
            }
        }

        public override int Dy2
        {
            get
            {
                return (int)(Dy1 + t2d.ext.cy);
            }
            set 
            {
                t2d.ext.cy = (value - Dy1);
            }
        }


        public override int Dx2
        {
            get
            {
                return (int)(Dx1 + t2d.ext.cx);
            }
            set 
            {
                t2d.ext.cx = (value - Dx1);
            }
        }
    }



}