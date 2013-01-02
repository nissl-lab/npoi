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
        long GetId();

        /**
         * @return axis position
         */
        AxisPosition GetPosition();

        /**
         * @param position new axis position
         */
        void SetPosition(AxisPosition position);

        /**
         * @return axis number format
         */
        String GetNumberFormat();

        /**
         * @param format axis number format
         */
        void SetNumberFormat(String format);

        /**
         * @return true if log base is defined, false otherwise
         */
        bool IsSetLogBase();

        /**
         * @param logBase a number between 2 and 1000 (inclusive)
         * @throws ArgumentException if log base not within allowed range
         */
        void SetLogBase(double logBase);

        /**
         * @return axis log base or 0.0 if not Set
         */
        double GetLogBase();

        /**
         * @return true if minimum value is defined, false otherwise
         */
        bool IsSetMinimum();

        /**
         * @return axis minimum or 0.0 if not Set
         */
        double GetMinimum();

        /**
         * @param min axis minimum
         */
        void SetMinimum(double min);

        /**
         * @return true if maximum value is defined, false otherwise
         */
        bool IsSetMaximum();

        /**
         * @return axis maximum or 0.0 if not Set
         */
        double GetMaximum();

        /**
         * @param max axis maximum
         */
        void SetMaximum(double max);

        /**
         * @return axis orientation
         */
        AxisOrientation GetOrientation();

        /**
         * @param orientation axis orientation
         */
        void SetOrientation(AxisOrientation orientation);

        /**
         * @param crosses axis cross type
         */
        void SetCrosses(AxisCrosses crosses);

        /**
         * @return axis cross type
         */
        AxisCrosses GetCrosses();

        /**
         * Declare this axis cross another axis.
         * @param axis that this axis should cross
         */
        void CrossAxis(IChartAxis axis);
    }


}