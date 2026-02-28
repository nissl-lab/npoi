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
    /// Specifies whether to layout the plot area by its inside (not including axis
    /// and axis labels) or outside (including axis and axis labels).
    /// </summary>
    /// <remarks>
    /// @author Roman Kashitsyn
    /// </remarks>
    public enum LayoutTarget
    {
        /// <summary>
        /// Specifies that the plot area size shall determine the size of the plot area, not including the tick marks and axis labels.
        /// </summary>
        Inner,
        /// <summary>
        ///Specifies that the plot area size shall determine the
        ///size of the plot area, the tick marks, and the axis
        ///labels.
        /// </summary>
        Outer
    }


}