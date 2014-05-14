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

using NPOI.SS.UserModel.Charts;
namespace NPOI.XSSF.UserModel.Charts
{

    /**
     * @author Roman Kashitsyn
     */
    public class XSSFChartDataFactory : IChartDataFactory
    {

        private static XSSFChartDataFactory instance;

        private XSSFChartDataFactory()
            : base()
        {

        }

        /**
         * @return new scatter chart data instance
         */
        public IScatterChartData<Tx, Ty> CreateScatterChartData<Tx, Ty>()
        {
            return new XSSFScatterChartData<Tx, Ty>();
        }

        public ILineChartData<Tx, Ty> CreateLineChartData<Tx, Ty>()
        {
            return new XSSFLineChartData<Tx, Ty>();
        }
        /**
         * @return factory instance
         */
        public static XSSFChartDataFactory GetInstance()
        {
            if (instance == null)
            {
                instance = new XSSFChartDataFactory();
            }
            return instance;
        }

    }

}