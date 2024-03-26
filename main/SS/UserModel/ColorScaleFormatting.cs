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

    using NPOI.HSSF.Record.CF;

    /// <summary>
    /// High level representation for the Color Scale / Colour Scale /
    ///  Color Gradient Formatting component of Conditional Formatting Settings
    /// </summary>
    public interface IColorScaleFormatting
    {

        /// <summary>
        /// <para>
        /// get or sets the number of control points to use to map
        ///  the colours. Should normally be 2 or 3.
        /// </para>
        /// <para>
        /// After updating, you need to ensure that the
        ///  <see cref="Threshold"/> count and Color count match
        /// </para>
        /// </summary>
        int NumControlPoints { get; set; }


        /// <summary>
        /// get or sets the list of colours that are interpolated
        ///  between.The number must match <see cref="NumControlPoints" />
        /// </summary>
        IColor[] Colors { get; set; }
        /// <summary>
        /// Gets the list of thresholds
        /// </summary>
        IConditionalFormattingThreshold[] Thresholds { get; set; }

        /// <summary>
        /// Creates a new, empty Threshold
        /// </summary>
        IConditionalFormattingThreshold CreateThreshold();
    }

}