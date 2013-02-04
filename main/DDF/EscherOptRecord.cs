
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

        public override short Instance
        {
            get
            {
                Instance = ((short)properties.Count);
                return base.Instance;
            }
        }

        /// <summary>
        /// Automatically recalculate the correct option
        /// </summary>
        /// <value></value>
        internal override short Options
        {
            get
            {
                // update values
                short tmp;
                tmp = Instance;
                tmp = Version;
                //base.Options = (short)((properties.Count<< 4) | 0x3);
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

        public override short Version
        {
            get
            {
                Version = 0x3;
                return base.Version;
            }
            set
            {
                if (value != 0x3)
                    throw new ArgumentException(RECORD_DESCRIPTION
                            + " can have only '0x3' version");

                base.Version = (value);
            }
        }


        public override String ToXml(String tab)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(tab).Append(FormatXmlRecordHeader(GetType().Name, HexDump.ToHex(RecordId), HexDump.ToHex(Version), HexDump.ToHex(Instance)));
            foreach (EscherProperty property in EscherProperties)
            {
                builder.Append(property.ToXml(tab + "\t"));
            }
            builder.Append(tab).Append("</").Append(GetType().Name).Append(">\n");
            return builder.ToString();
        }
    }
}
