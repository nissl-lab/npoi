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
using System;
using NPOI.HWPF.Model;
using NPOI.Util;
namespace NPOI.HWPF.UserModel
{

    /**
     * TODO: document me
     * 
     * @author Sergey Vladimirov (vlsergey {at} gmail {dot} com)
     */

    public class FieldImpl : Field
    {
        private PlexOfField endPlex;
        private PlexOfField separatorPlex;
        private PlexOfField startPlex;

        public FieldImpl(PlexOfField startPlex, PlexOfField separatorPlex,
                PlexOfField endPlex)
        {
            if (startPlex == null)
                throw new ArgumentException("startPlex == null");
            if (endPlex == null)
                throw new ArgumentException("endPlex == null");

            if (startPlex.Fld.GetBoundaryType() != FieldDescriptor.FIELD_BEGIN_MARK)
                throw new ArgumentException("startPlex (" + startPlex
                        + ") is not type of FIELD_BEGIN");

            if (separatorPlex != null
                    && separatorPlex.Fld.GetBoundaryType() != FieldDescriptor.FIELD_SEPARATOR_MARK)
                throw new ArgumentException("separatorPlex" + separatorPlex
                        + ") is not type of FIELD_SEPARATOR");

            if (endPlex.Fld.GetBoundaryType() != FieldDescriptor.FIELD_END_MARK)
                throw new ArgumentException("endPlex (" + endPlex
                        + ") is not type of FIELD_END");

            this.startPlex = startPlex;
            this.separatorPlex = separatorPlex;
            this.endPlex = endPlex;
        }


        public class FieldSubRange : Range
        {
            string className;

            public FieldSubRange(int start, int end, Range parent, string className)
                : base(start, end, parent)
            {
                this.className = className;
            }
            public override String ToString()
            {
                return this.className + " (" + base.ToString() + ")";
            }
        }

        public Range FirstSubrange(Range parent)
        {
            if (HasSeparator())
            {
                if (GetMarkStartOffset() + 1 == GetMarkSeparatorOffset())
                    return null;

                return new FieldSubRange(GetMarkStartOffset() + 1,
                        GetMarkSeparatorOffset(), parent, "FieldSubrange1");

            }

            if (GetMarkStartOffset() + 1 == GetMarkEndOffset())
                return null;

            return new FieldSubRange(GetMarkStartOffset() + 1, GetMarkEndOffset(), parent, "FieldSubrange1");

        }

        /**
         * @return character position of first character after field (i.e.
         *         {@link #getMarkEndOffset()} + 1)
         */
        public int GetFieldEndOffset()
        {
            /*
             * sometimes plex looks like [100, 2000), where 100 is the position of
             * field-end character, and 2000 - some other char position, far away
             * from field (not inside). So taking into account only start --sergey
             */
            return endPlex.FcStart + 1;
        }

        /**
         * @return character position of first character in field (i.e.
         *         {@link #getFieldStartOffset()})
         */
        public int GetFieldStartOffset()
        {
            return startPlex.FcStart;
        }

        public CharacterRun GetMarkEndCharacterRun(Range parent)
        {
            return new Range(GetMarkEndOffset(), GetMarkEndOffset() + 1, parent)
                    .GetCharacterRun(0);
        }

        /**
         * @return character position of end field mark
         */
        public int GetMarkEndOffset()
        {
            return endPlex.FcStart;
        }

        public CharacterRun GetMarkSeparatorCharacterRun(Range parent)
        {
            if (!HasSeparator())
                return null;

            return new Range(GetMarkSeparatorOffset(),
                    GetMarkSeparatorOffset() + 1, parent).GetCharacterRun(0);
        }

        /**
         * @return character position of separator field mark (if present,
         *         {@link NullPointerException} otherwise)
         */
        public int GetMarkSeparatorOffset()
        {
            return separatorPlex.FcStart;
        }

        public CharacterRun GetMarkStartCharacterRun(Range parent)
        {
            return new Range(GetMarkStartOffset(), GetMarkStartOffset() + 1,
                    parent).GetCharacterRun(0);
        }

        /**
         * @return character position of start field mark
         */
        public int GetMarkStartOffset()
        {
            return startPlex.FcStart;
        }

        public int Type
        {
            get
            {
                return startPlex.Fld.GetFieldType();
            }
        }

        public bool HasSeparator()
        {
            return separatorPlex != null;
        }

        public bool IsHasSep()
        {
            return endPlex.Fld.IsFHasSep();
        }

        public bool IsLocked()
        {
            return endPlex.Fld.IsFLocked();
        }

        public bool IsNested()
        {
            return endPlex.Fld.IsFNested();
        }

        public bool IsPrivateResult()
        {
            return endPlex.Fld.IsFPrivateResult();
        }

        public bool IsResultDirty()
        {
            return endPlex.Fld.IsFResultDirty();
        }

        public bool IsResultEdited()
        {
            return endPlex.Fld.IsFResultEdited();
        }

        public bool IsZombieEmbed()
        {
            return endPlex.Fld.IsFZombieEmbed();
        }

        public Range SecondSubrange(Range parent)
        {
            if (!HasSeparator()
                    || GetMarkSeparatorOffset() + 1 == GetMarkEndOffset())
                return null;

            return new FieldSubRange(GetMarkSeparatorOffset() + 1, GetMarkEndOffset(),
                    parent, "FieldSubrange2");
        }


        public override String ToString()
        {
            return "Field [" + GetFieldStartOffset() + "; " + GetFieldEndOffset()
                    + "] (type: 0x" + StringUtil.ToHexString(Type) + " = "
                    + GetType() + " )";
        }
    }

}

