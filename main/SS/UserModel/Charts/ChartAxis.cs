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
namespace NPOI.SS.UserModel.Charts
{

    /// <summary>
    /// High level representation of chart axis.
    /// </summary>
    /// @author Roman Kashitsyn
    public interface IChartAxis
    {

        /// <summary>
        /// </summary>
        /// <returns>axis id</returns>
        long Id { get; }

        /// <summary>
        /// get or set axis position
        /// </summary>
        AxisPosition Position { get; set; }

        /// <summary>
        /// get or set axis number format
        /// </summary>
        String NumberFormat { get; set; }

        /// <summary>
        /// </summary>
        /// <returns>true if log base is defined, false otherwise</returns>
        bool IsSetLogBase { get; }

        /// <summary>
        /// </summary>
        /// <param name="logBase">a number between 2 and 1000 (inclusive)</param>
        /// <returns>axis log base or 0.0 if not Set</returns>
        /// <exception cref="ArgumentException">if log base not within allowed range</exception>
        double LogBase { get; set; }

        /// <summary>
        /// </summary>
        /// <returns>true if minimum value is defined, false otherwise</returns>
        bool IsSetMinimum { get; }

        /// <summary>
        /// get or set axis minimum
        /// 0.0 if not Set
        /// </summary>
        double Minimum { get; set; }

        /// <summary>
        /// </summary>
        /// <returns>true if maximum value is defined, false otherwise</returns>
        bool IsSetMaximum { get; }

        /// <summary>
        /// get or set axis maximum
        /// 0.0 if not Set
        /// </summary>
        double Maximum { get; set; }

        /// <summary>
        /// get or set axis orientation
        /// </summary>
        AxisOrientation Orientation { get; set; }

        /// <summary>
        /// get or set axis cross type
        /// </summary>
        AxisCrosses Crosses { get; set; }

        /// <summary>
        /// Declare this axis cross another axis.
        /// </summary>
        /// <param name="axis">that this axis should cross</param>
        void CrossAxis(IChartAxis axis);

        /// <summary>
        /// </summary>
        /// <returns>visibility of the axis.</returns>
        bool IsVisible { get; set; }

        /// <summary>
        /// </summary>
        /// <returns>major tick mark.</returns>
        AxisTickMark MajorTickMark { get; set; }

        /// <summary>
        /// </summary>
        /// <returns>minor tick mark.</returns>
        AxisTickMark MinorTickMark { get; set; }
    }


}