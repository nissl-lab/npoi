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

namespace NPOI.HSSF.UserModel
{
    using System;

    /// <summary>
    /// @author Glen Stampoultzis  (glens at baselinksoftware.com)
    /// </summary>
    public class HSSFPolygon : HSSFShape
    {
        int[] xPoints;
        int[] yPoints;
        int drawAreaWidth = 100;
        int drawAreaHeight = 100;

        public HSSFPolygon(HSSFShape parent, HSSFAnchor anchor)
            : base(parent, anchor)
        {

        }

        public int[] XPoints
        {
            get { return xPoints; }
            set { this.xPoints = CloneArray(value); }
        }

        public int[] YPoints
        {
            get
            {
                return yPoints;
            }
            set { this.yPoints = CloneArray(value); }
        }

        public void SetPoints(int[] xPoints, int[] yPoints)
        {
            this.xPoints = CloneArray(xPoints);
            this.yPoints = CloneArray(yPoints);
        }

        private int[] CloneArray(int[] a)
        {
            int[] result = new int[a.Length];
            for (int i = 0; i < a.Length; i++)
                result[i] = a[i];

            return result;
        }

        /**
         * Defines the width and height of the points in the polygon
         * @param width
         * @param height
         */
        public void SetPolygonDrawArea(int width, int height)
        {
            this.drawAreaWidth = width;
            this.drawAreaHeight = height;
        }

        public int DrawAreaWidth
        {
            get
            {
                return drawAreaWidth;
            }
            set { this.drawAreaWidth = value; }
        }

        public int DrawAreaHeight
        {
            get
            {
                return drawAreaHeight;
            }
            set { this.drawAreaHeight = value; }
        }


    }
}