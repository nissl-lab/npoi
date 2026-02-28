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
    /// High level representation of chart legend.
    /// </summary>
    /// <remarks>@author Roman Kashitsyn</remarks>
    public interface IChartLegend : ManuallyPositionable
    {
        /// <summary>
        /// legend position
        /// </summary>
        /// <returns></returns>
        LegendPosition Position
        {
            get;
            set;
        }

        /// <summary>
        /// If true the legend is positioned over the chart area otherwise
        /// the legend is displayed next to it.
        /// Default is no overlay.
        /// </summary>
        bool IsOverlay
        {
            get;
            set;
        }
    }


}