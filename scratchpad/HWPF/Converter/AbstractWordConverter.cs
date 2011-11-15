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
using System.Collections.Generic;
using System.Text;
using NPOI.HWPF.UserModel;

namespace NPOI.HWPF.Converter
{
    public class AbstractWordConverter : AbstractWordUtils
    {
        private class Structure : IComparable<Structure>
        {
            int end;
            int start;
            Object structure;

            Structure(Bookmark bookmark)
            {
                this.start = bookmark.GetStart();
                this.end = bookmark.GetEnd();
                this.structure = bookmark;
            }

            Structure(Field field)
            {
                this.start = field.GetFieldStartOffset();
                this.end = field.GetFieldEndOffset();
                this.structure = field;
            }

            public int compareTo(Structure o)
            {
                return start < o.start ? -1 : start == o.start ? 0 : 1;
            }

            public override String ToString()
            {
                return "Structure [" + start + "; " + end + "): "
                        + structure.ToString();
            }

            #region IComparable<Structure> 成员

            public int CompareTo(Structure other)
            {
                return start < other.start ? -1 : start == other.start ? 0 : 1;
            }

            #endregion
        }
    }
    
}
