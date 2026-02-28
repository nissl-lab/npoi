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

namespace NPOI.SS.UserModel.Charts
{

    /**
     * High level representation of chart element manual layout.
     *
     * @author Roman Kashitsyn
     */
    public interface IManualLayout
    {

        /**
         * Sets the layout target.
         * @param target new layout target.
         */
        void SetTarget(LayoutTarget target);

        /**
         * Returns current layout target.
         * @return current layout target
         */
        LayoutTarget GetTarget();

        /**
         * Sets the x-coordinate layout mode.
         * @param mode new x-coordinate layout mode.
         */
        void SetXMode(LayoutMode mode);

        /**
         * Returns current x-coordinnate layout mode.
         * @return current x-coordinate layout mode.
         */
        LayoutMode GetXMode();

        /**
         * Sets the y-coordinate layout mode.
         * @param mode new y-coordinate layout mode.
         */
        void SetYMode(LayoutMode mode);

        /**
         * Returns current y-coordinate layout mode.
         * @return current y-coordinate layout mode.
         */
        LayoutMode GetYMode();

        /**
         * Returns the x location of the chart element.
         * @return the x location (left) of the chart element or 0.0 if
         *         not Set.
         */
        double GetX();

        /**
         * Specifies the x location (left) of the chart element as a
         * fraction of the width of the chart. If Left Mode is Factor,
         * then the position is relative to the default position for the
         * chart element.
         */
        void SetX(double x);


        /**
         * Returns current y location of the chart element.
         * @return the y location (top) of the chart element or 0.0 if not
         *         Set.
         */
        double GetY();

        /**
         * Specifies the y location (top) of the chart element as a
         * fraction of the height of the chart. If Top Mode is Factor,
         * then the position is relative to the default position for the
         * chart element.
         */
        void SetY(double y);


        /**
         * Specifies how to interpret the Width element for this manual
         * layout.
         * @param mode new width layout mode of this manual layout.
         */
        void SetWidthMode(LayoutMode mode);


        /**
         * Returns current width mode of this manual layout.
         * @return width mode of this manual layout.
         */
        LayoutMode GetWidthMode();

        /**
         * Specifies how to interpret the Height element for this manual
         * layout.
         * @param mode new height mode of this manual layout.
         */
        void SetHeightMode(LayoutMode mode);

        /**
         * Returns current height mode of this 
         * @return height mode of this manual layout.
         */
        LayoutMode GetHeightMode();

        /**
         * Specifies the width (if Width Mode is Factor) or right (if
         * Width Mode is Edge) of the chart element as a fraction of the
         * width of the chart.
         * @param ratio a fraction of the width of the chart.
         */
        void SetWidthRatio(double ratio);

        /**
         * Returns current fraction of the width of the chart.
         * @return fraction of the width of the chart or 0.0 if not Set.
         */
        double GetWidthRatio();

        /**
         * Specifies the height (if Height Mode is Factor) or bottom (if
         * Height Mode is edge) of the chart element as a fraction of the
         * height of the chart.
         * @param ratio a fraction of the height of the chart.
         */
        void SetHeightRatio(double ratio);

        /**
         * Returns current fraction of the height of the chart.
         * @return fraction of the height of the chart or 0.0 if not Set.
         */
        double GetHeightRatio();

    }
}


