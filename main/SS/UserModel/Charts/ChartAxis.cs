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

    /**
     * High level representation of chart axis.
     *
     * @author Roman Kashitsyn
     */
    public interface IChartAxis
    {

        /**
         * @return axis id
         */
        long Id { get; }

        /**
         * get or set axis position
         */
        AxisPosition Position { get; set; }

        /**
         * get or set axis number format
         */
        String NumberFormat { get; set; }

        /**
         * @return true if log base is defined, false otherwise
         */
        bool IsSetLogBase { get; }

        /**
         * @param logBase a number between 2 and 1000 (inclusive)
         * @return axis log base or 0.0 if not Set
         * @throws ArgumentException if log base not within allowed range
         */
        double LogBase { get; set; }

        /**
         * @return true if minimum value is defined, false otherwise
         */
        bool IsSetMinimum { get; }

        /**
         * get or set axis minimum 
         * 0.0 if not Set
         */
        double Minimum { get; set; }

        /**
         * @return true if maximum value is defined, false otherwise
         */
        bool IsSetMaximum { get; }

        /**
         * get or set axis maximum 
         * 0.0 if not Set
         */
        double Maximum { get; set; }

        /**
         * get or set axis orientation
         */
        AxisOrientation Orientation { get; set; }

        /**
         * get or set axis cross type
         */
        AxisCrosses Crosses { get; set; }

        /**
         * Declare this axis cross another axis.
         * @param axis that this axis should cross
         */
        void CrossAxis(IChartAxis axis);

        /**
         * @return visibility of the axis.
         */
        bool IsVisible { get; set; }

        /**
         * @return major tick mark.
         */
        AxisTickMark MajorTickMark { get; set; }

        /**
         * @return minor tick mark.
         */
        AxisTickMark MinorTickMark { get; set; }
    }


}