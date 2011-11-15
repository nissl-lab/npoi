using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HWPF.UserModel
{
    interface Field
    {
        Range firstSubrange(Range parent);

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

        CharacterRun getMarkEndCharacterRun(Range parent);

        /**
         * @return character position of end field mark
         */
        int getMarkEndOffset();

        CharacterRun getMarkSeparatorCharacterRun(Range parent);

        /**
         * @return character position of separator field mark (if present,
         *         {@link NullPointerException} otherwise)
         */
        int getMarkSeparatorOffset();

        CharacterRun getMarkStartCharacterRun(Range parent);

        /**
         * @return character position of start field mark
         */
        int getMarkStartOffset();

        int getType();

        bool hasSeparator();

        bool isHasSep();

        bool isLocked();

        bool isNested();

        bool isPrivateResult();

        bool isResultDirty();

        bool isResultEdited();

        bool isZombieEmbed();

        Range secondSubrange(Range parent);
    }
}
