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

namespace NPOI.HWPF.UserModel
{
    public interface Field
    {
        Range FirstSubrange(Range parent);

        /**
         * @return character position of first character after field (i.e.
         *         {@link #getMarkEndOffset()} + 1)
         */
        int GetFieldEndOffset();

        /**
         * @return character position of first character in field (i.e.
         *         {@link #getFieldStartOffset()})
         */
        int GetFieldStartOffset();

        CharacterRun GetMarkEndCharacterRun(Range parent);

        /**
         * @return character position of end field mark
         */
        int GetMarkEndOffset();

        CharacterRun GetMarkSeparatorCharacterRun(Range parent);

        /**
         * @return character position of separator field mark (if present,
         *         {@link NullPointerException} otherwise)
         */
        int GetMarkSeparatorOffset();

        CharacterRun GetMarkStartCharacterRun(Range parent);

        /**
         * @return character position of start field mark
         */
        int GetMarkStartOffset();

        int Type { get; }

        bool HasSeparator();

        bool IsHasSep();

        bool IsLocked();

        bool IsNested();

        bool IsPrivateResult();

        bool IsResultDirty();

        bool IsResultEdited();

        bool IsZombieEmbed();

        Range SecondSubrange(Range parent);
    }
}
