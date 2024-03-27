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

    /// <summary>
    /// High level representation of chart element manual layout.
    /// </summary>
    /// @author Roman Kashitsyn
    public interface IManualLayout
    {

        /// <summary>
        /// Sets the layout target.
        /// </summary>
        /// <param name="target">new layout target.</param>
        void SetTarget(LayoutTarget target);

        /// <summary>
        /// Returns current layout target.
        /// </summary>
        /// <returns>current layout target</returns>
        LayoutTarget GetTarget();

        /// <summary>
        /// Sets the x-coordinate layout mode.
        /// </summary>
        /// <param name="mode">new x-coordinate layout mode.</param>
        void SetXMode(LayoutMode mode);

        /// <summary>
        /// Returns current x-coordinnate layout mode.
        /// </summary>
        /// <returns>current x-coordinate layout mode.</returns>
        LayoutMode GetXMode();

        /// <summary>
        /// Sets the y-coordinate layout mode.
        /// </summary>
        /// <param name="mode">new y-coordinate layout mode.</param>
        void SetYMode(LayoutMode mode);

        /// <summary>
        /// Returns current y-coordinate layout mode.
        /// </summary>
        /// <returns>current y-coordinate layout mode.</returns>
        LayoutMode GetYMode();

        /// <summary>
        /// Returns the x location of the chart element.
        /// </summary>
        /// <returns>the x location (left) of the chart element or 0.0 if
        /// not Set.
        /// </returns>
        double GetX();

        /// <summary>
        /// Specifies the x location (left) of the chart element as a
        /// fraction of the width of the chart. If Left Mode is Factor,
        /// then the position is relative to the default position for the
        /// chart element.
        /// </summary>
        void SetX(double x);


        /// <summary>
        /// Returns current y location of the chart element.
        /// </summary>
        /// <returns>the y location (top) of the chart element or 0.0 if not
        /// Set.
        /// </returns>
        double GetY();

        /// <summary>
        /// Specifies the y location (top) of the chart element as a
        /// fraction of the height of the chart. If Top Mode is Factor,
        /// then the position is relative to the default position for the
        /// chart element.
        /// </summary>
        void SetY(double y);


        /// <summary>
        /// Specifies how to interpret the Width element for this manual
        /// layout.
        /// </summary>
        /// <param name="mode">new width layout mode of this manual layout.</param>
        void SetWidthMode(LayoutMode mode);


        /// <summary>
        /// Returns current width mode of this manual layout.
        /// </summary>
        /// <returns>width mode of this manual layout.</returns>
        LayoutMode GetWidthMode();

        /// <summary>
        /// Specifies how to interpret the Height element for this manual
        /// layout.
        /// </summary>
        /// <param name="mode">new height mode of this manual layout.</param>
        void SetHeightMode(LayoutMode mode);

        /// <summary>
        /// Returns current height mode of this
        /// </summary>
        /// <returns>height mode of this manual layout.</returns>
        LayoutMode GetHeightMode();

        /// <summary>
        /// Specifies the width (if Width Mode is Factor) or right (if
        /// Width Mode is Edge) of the chart element as a fraction of the
        /// width of the chart.
        /// </summary>
        /// <param name="ratio">a fraction of the width of the chart.</param>
        void SetWidthRatio(double ratio);

        /// <summary>
        /// Returns current fraction of the width of the chart.
        /// </summary>
        /// <returns>fraction of the width of the chart or 0.0 if not Set.</returns>
        double GetWidthRatio();

        /// <summary>
        /// Specifies the height (if Height Mode is Factor) or bottom (if
        /// Height Mode is edge) of the chart element as a fraction of the
        /// height of the chart.
        /// </summary>
        /// <param name="ratio">a fraction of the height of the chart.</param>
        void SetHeightRatio(double ratio);

        /// <summary>
        /// Returns current fraction of the height of the chart.
        /// </summary>
        /// <returns>fraction of the height of the chart or 0.0 if not Set.</returns>
        double GetHeightRatio();

    }
}


