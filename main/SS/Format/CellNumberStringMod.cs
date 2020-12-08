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
namespace NPOI.SS.Format
{
    using System;

    /**
* Internal helper class for CellNumberFormatter
*
* This class represents a single modification to a result string.  The way
* this works is complicated, but so is numeric formatting.  In general, for
* most formats, we use a DecimalFormat object that will Put the string out
* in a known format, usually with all possible leading and trailing zeros.
* We then walk through the result and the original format, and note any
* modifications that need to be made.  Finally, we go through and apply
* them all, dealing with overlapping modifications.
*/

    public class CellNumberStringMod : IComparable<CellNumberStringMod>
    {
        public const int BEFORE = 1;
        public const int AFTER = 2;
        public const int REPLACE = 3;

        private CellNumberFormatter.Special special;
        private int op;
        //private CharSequence toAdd;
        private string toAdd;
        private CellNumberFormatter.Special end;
        private bool startInclusive;
        private bool endInclusive;

        public CellNumberStringMod(CellNumberFormatter.Special special, string toAdd, int op)
        {
            this.special = special;
            this.toAdd = toAdd;
            this.op = op;
        }

        public CellNumberStringMod(CellNumberFormatter.Special start, bool startInclusive, CellNumberFormatter.Special end, bool endInclusive, char toAdd)
                : this(start, startInclusive, end, endInclusive)
        {
            ;
            this.toAdd = toAdd + "";
        }

        public CellNumberStringMod(CellNumberFormatter.Special start, bool startInclusive, CellNumberFormatter.Special end, bool endInclusive)
        {
            special = start;
            this.startInclusive = startInclusive;
            this.end = end;
            this.endInclusive = endInclusive;
            op = REPLACE;
            toAdd = "";
        }

        public int CompareTo(CellNumberStringMod that)
        {
            int diff = special.pos - that.special.pos;
            return (diff != 0) ? diff : (op - that.op);
        }


        public override bool Equals(Object that)
        {
            try
            {
                return CompareTo((CellNumberStringMod)that) == 0;
            }
            catch (Exception)
            {
                // NullPointerException or CastException
                return false;
            }
        }


        public override int GetHashCode()
        {
            return special.GetHashCode() + op;
        }

        public CellNumberFormatter.Special GetSpecial()
        {
            return special;
        }

        public int Op
        {
            get
            {
                return op;
            }
        }

        public string ToAdd
        {
            get
            {
                return toAdd;
            }
        }

        public CellNumberFormatter.Special End
        {
            get
            {
                return end;
            }
        }

        public bool IsStartInclusive
        {
            get
            {
                return startInclusive;
            }
        }

        public bool IsEndInclusive
        {
            get
            {
                return endInclusive;
            }
        }
    }

}