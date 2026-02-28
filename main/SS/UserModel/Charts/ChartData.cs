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
    /// A base for all chart data types.
    /// </summary>
    /// <remarks>
    /// @author  Roman Kashitsyn
    /// </remarks>
    public interface IChartData
    {
        /// <summary>
        /// Fills a chart with data specified by implementation.
        /// </summary>
        /// <param name="chart">a chart to fill in</param>
        /// <param name="axis">chart axis to use</param>
        void FillChart(IChart chart, params IChartAxis[] axis);
    }

}
