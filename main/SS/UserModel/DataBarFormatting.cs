/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for Additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

namespace NPOI.SS.UserModel
{
    using System;

    /**
     * High level representation for the DataBar Formatting 
     *  component of Conditional Formatting Settings
     */
    public interface IDataBarFormatting
    {
        /**
         * Is the bar Drawn from Left-to-Right, or from
         *  Right-to-Left
         */
        bool IsLeftToRight { get; set; }
        /**
         * Should Icon + Value be displayed, or only the Icon?
         */
        bool IsIconOnly { get; set; }
        /**
         * How much of the cell width, in %, should be given to
         *  the min value?
         */
        int WidthMin { get; set; }
        /**
         * How much of the cell width, in %, should be given to
         *  the max value?
         */
        int WidthMax { get; set; }

        IColor Color { get; set; }

        /**
         * The threshold that defines "everything from here down is minimum"
         */
        IConditionalFormattingThreshold MinThreshold { get; }
        /**
         * The threshold that defines "everything from here up is maximum"
         */
        IConditionalFormattingThreshold MaxThreshold { get; }
    }

}