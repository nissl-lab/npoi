/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.SS.UserModel
{
    /// <summary>
    /// The enumeration value indicating which borders to Draw in a Property Template
    /// </summary>
    public enum BorderExtent
    {
        /// <summary>
        /// No properties defined. This can be used to remove existing properties
        /// from the PropertyTemplate.
        /// </summary>
        NONE,

        /// <summary>
        /// All borders, that is top, bottom, left and right, including interior
        /// borders for the range. Does not include diagonals which are different
        /// and not implemented here.
        /// </summary>
        ALL,

        /// <summary>
        /// All inside borders. This is top, bottom, left, and right borders, but
        /// restricted to the interior borders for the range. For a range of one
        /// cell, this will produce no borders.
        /// </summary>
        INSIDE,

        /// <summary>
        /// All outside borders. That is top, bottom, left and right borders that
        /// bound the range only.
        /// </summary>
        OUTSIDE,

        /// <summary>
        /// This is just the top border for the range. No interior borders will
        /// be produced.
        /// </summary>
        TOP,

        /// <summary>
        /// This is just the bottom border for the range. No interior borders
        /// will be produced.
        /// </summary>
        BOTTOM,

        /// <summary>
        /// This is just the left border for the range, no interior borders will
        /// be produced.
        /// </summary>
        LEFT,

        /// <summary>
        /// This is just the right border for the range, no interior borders will
        /// be produced.
        /// </summary>
        RIGHT,

        /// <summary>
        /// This is all horizontal borders for the range, including interior and
        /// outside borders.
        /// </summary>
        HORIZONTAL,

        /// <summary>
        /// This is just the interior horizontal borders for the range.
        /// </summary>
        INSIDE_HORIZONTAL,

        /// <summary>
        /// This is just the outside horizontal borders for the range.
        /// </summary>
        OUTSIDE_HORIZONTAL,

        /// <summary>
        /// This is all vertical borders for the range, including interior and
        /// outside borders.
        /// </summary>
        VERTICAL,

        /// <summary>
        /// This is just the interior vertical borders for the range.
        /// </summary>
        INSIDE_VERTICAL,

        /// <summary>
        /// This is just the outside vertical borders for the range.
        /// </summary>
        OUTSIDE_VERTICAL
    }
}
