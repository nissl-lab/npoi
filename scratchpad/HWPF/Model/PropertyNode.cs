
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License Is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HWPF.Model
{
    using System;
    using System.Collections;
    using NPOI.Util;
using System.Collections.Generic;

    /**
     * Represents a lightweight node in the Trees used to store content
     *  properties.
     * This only ever works in characters. For the few odd cases when
     *  the start and end aren't in characters (eg PAPX and CHPX), use
     *  {@link BytePropertyNode} between you and this.
     *
     * @author Ryan Ackley
     */
    public abstract class PropertyNode : IComparable
    {
            public class EndComparator:IComparer<PropertyNode>
            {
                public static EndComparator instance = new EndComparator();

                public int Compare( PropertyNode o1, PropertyNode o2 )
                {
                    int thisVal = o1.End;
                    int anotherVal = o2.End;
                    return ( thisVal < anotherVal ? -1 : ( thisVal == anotherVal ? 0
                            : 1 ) );
                }
            }

            public class StartComparator:IComparer<PropertyNode>
            {
                public static StartComparator instance = new StartComparator();

                public int Compare(PropertyNode o1, PropertyNode o2)
                {
                    int thisVal = o1.Start;
                    int anotherVal = o2.Start;
                    return ( thisVal < anotherVal ? -1 : ( thisVal == anotherVal ? 0
                            : 1 ) );
                }
            }

            public class CHPXComparator : IComparer<CHPX>
            {
                public static CHPXComparator instance = new CHPXComparator();

                public int Compare(CHPX o1, CHPX o2)
                {
                    int thisVal = o1.Start;
                    int anotherVal = o2.Start;
                    return (thisVal < anotherVal ? -1 : (thisVal == anotherVal ? 0
                            : 1));
                } 
            }

            public class PAPXComparator : IComparer<PAPX>
            {
                public static PAPXComparator instance = new PAPXComparator();

                public int Compare(PAPX o1, PAPX o2)
                {
                    int thisVal = o1.Start;
                    int anotherVal = o2.Start;
                    return (thisVal < anotherVal ? -1 : (thisVal == anotherVal ? 0
                            : 1));
                }
            }

            public class SEPXComparator : IComparer<SEPX>
            {
                public static SEPXComparator instance = new SEPXComparator();

                public int Compare(SEPX o1, SEPX o2)
                {
                    int thisVal = o1.Start;
                    int anotherVal = o2.Start;
                    return (thisVal < anotherVal ? -1 : (thisVal == anotherVal ? 0
                            : 1));
                }
            }

        protected Object _buf;
        /** The start, in characters */
        private int _cpStart;
        /** The end, in characters */
        private int _cpEnd;


        /**
         * @param fcStart The start of the text for this property, in characters.
         * @param fcEnd The end of the text for this property, in characters.
         * @param buf FIXME: Old documentation Is: "grpprl The property description in compressed form."
         */
        protected PropertyNode(int fcStart, int fcEnd, Object buf)
        {
            _cpStart = fcStart;
            _cpEnd = fcEnd;
            _buf = buf;

            if (_cpStart < 0)
            {
                Console.WriteLine("A property claimed to start before zero, at " + _cpStart + "! Resetting it to zero, and hoping for the best");
                _cpStart = 0;
            }
            if (_cpEnd < _cpStart)
            {
                Console.WriteLine("A property claimed to end (" + _cpEnd
                        + ") before start! "
                        + "Resetting end to start, and hoping for the best");
                _cpEnd = _cpStart;
            }
        }

        /**
         * @return The start offset of this property's text.
         */
        public int Start
        {
            get
            {
                return _cpStart;
            }
            set
            {
                _cpStart = value;
            }
        }


        /**
         * @return The offset of the end of this property's text.
         */
        public int End
        {
            get
            {
                return _cpEnd;
            }
            set
            {
                _cpEnd = value;
            }
        }

        /**
         * Adjust for a deletion that can span multiple PropertyNodes.
         * @param start
         * @param length
         */
        public virtual void AdjustForDelete(int start, int length)
        {
            int end = start + length;

            if (_cpEnd > start)
            {
                // The start of the change Is before we end

                if (_cpStart < end)
                {
                    // The delete was somewhere in the middle of us
                    _cpEnd = end >= _cpEnd ? start : _cpEnd - length;
                    _cpStart = Math.Min(start, _cpStart);
                }
                else
                {
                    // The delete was before us
                    _cpEnd -= length;
                    _cpStart -= length;
                }
            }
        }

        protected bool LimitsAreEqual(Object o)
        {
            return ((PropertyNode)o).Start == _cpStart &&
                   ((PropertyNode)o).End == _cpEnd;

        }

        public override bool Equals(Object o)
        {
            if (LimitsAreEqual(o))
            {
                Object testBuf = ((PropertyNode)o)._buf;
                if (testBuf is byte[] && _buf is byte[])
                {
                    return Arrays.Equals((byte[])testBuf, (byte[])_buf);
                }
                return _buf.Equals(testBuf);
            }
            return false;
        }

        /**
         * Used for sorting in collections.
         */
        public int CompareTo(Object o)
        {
            int cpEnd = ((PropertyNode)o).End;
            if (_cpEnd == cpEnd)
            {
                return 0;
            }
            else if (_cpEnd < cpEnd)
            {
                return -1;
            }
            else
            {
                return 1;
            }
        }

        public override int GetHashCode()
        {
            return this._cpStart * 31 + this._buf.GetHashCode();
        }
    }
}