
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

namespace NPOI.DDF
{
    /// <summary>
    /// Defines the constants for the various possible shape paths.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    public class EscherShapePathProperty : EscherSimpleProperty
    {

        public const int LINE_OF_STRAIGHT_SEGMENTS = 0;
        public const int CLOSED_POLYGON = 1;
        public const int CURVES = 2;
        public const int CLOSED_CURVES = 3;
        public const int COMPLEX = 4;

        /// <summary>
        /// Initializes a new instance of the <see cref="EscherShapePathProperty"/> class.
        /// </summary>
        /// <param name="propertyNumber">The property number.</param>
        /// <param name="shapePath">The shape path.</param>
        public EscherShapePathProperty(short propertyNumber, int shapePath)
            : base(propertyNumber, false, false, shapePath)
        {

        }

    }

}