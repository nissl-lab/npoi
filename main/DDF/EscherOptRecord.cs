
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
    using System;
    using System.Text;
    using System.Collections;
    using System.Collections.Generic;
    using NPOI.Util;


    /// <summary>
    /// The opt record is used to store property values for a shape.  It is the key to determining
    /// the attributes of a shape.  Properties can be of two types: simple or complex.  Simple types
    /// are fixed Length.  Complex properties are variable Length.
    /// @author Glen Stampoultzis
    /// </summary>
    public class EscherOptRecord : AbstractEscherOptRecord
    {
        public const short RECORD_ID = unchecked((short)0xF00B);
        public const String RECORD_DESCRIPTION = "msofbtOPT";


        /// <summary>
        /// Automatically recalculate the correct option
        /// </summary>
        /// <value></value>
        public override short Options
        {
            get
            {
                base.Options = (short)((properties.Count<< 4) | 0x3);
                return base.Options;
            }
        }

        /// <summary>
        /// The short name for this record
        /// </summary>
        /// <value></value>
        public override String RecordName
        {
            get { return "Opt"; }
        }
    }
}
