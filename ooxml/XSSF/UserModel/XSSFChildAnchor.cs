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

using NPOI.OpenXmlFormats.Dml;
using System;
namespace NPOI.XSSF.UserModel
{

    /**
     * @author Yegor Kozlov
     */
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


        public CT_Transform2D GetCT_Transform2D()
        {
            return t2d;
        }

        public int GetDx1()
        {
            return (int)t2d.off.x;
        }

        public void SetDx1(int dx1)
        {
            t2d.off.y = (dx1);
        }

        public int GetDy1()
        {
            return (int)t2d.off.y;
        }

        public void SetDy1(int dy1)
        {
            t2d.off.y = (dy1);
        }

        public int GetDy2()
        {
            return (int)(Dy1 + t2d.ext.cy);
        }

        public void SetDy2(int dy2)
        {
            t2d.ext.cy = (dy2 - GetDy1());
        }

        public int GetDx2()
        {
            return (int)(Dx1 + t2d.ext.cx);
        }

        public void SetDx2(int dx2)
        {
            t2d.ext.cx = (dx2 - GetDx1());
        }
    }



}