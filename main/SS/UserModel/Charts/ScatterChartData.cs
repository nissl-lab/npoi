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

using System.Collections.Generic;
using NPOI.SS.Util;
namespace NPOI.SS.UserModel.Charts
{


    /**
     * Data for a Scatter Chart
     */
    public interface IScatterChartData<Tx,Ty> : IChartData
    {
        /**
         * @param xs data source to be used for X axis values
         * @param ys data source to be used for Y axis values
         * @return a new scatter charts series
         */
        IScatterChartSeries<Tx, Ty> AddSeries(IChartDataSource<Tx> xs, IChartDataSource<Ty> ys);

        /**
         * @return list of all series
         */
        List<IScatterChartSeries<Tx,Ty>> GetSeries();
    }


}