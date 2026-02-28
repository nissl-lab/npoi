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

using NPOI.Util;
using System;
using NPOI.HWPF.Model.Types;
namespace NPOI.HWPF.Model
{

    public class FieldDescriptor : FLDAbstractType
    {
        public const int FIELD_BEGIN_MARK = 0x13;
        public const int FIELD_SEPARATOR_MARK = 0x14;
        public const int FIELD_END_MARK = 0x15;

        public FieldDescriptor(byte[] data)
        {
            FillFields(data, 0);
        }

        public int GetBoundaryType()
        {
            return GetCh();
        }

        public int GetFieldType()
        {
            if (GetCh() != FIELD_BEGIN_MARK)
                throw new NotSupportedException(
                        "This field Is only defined for begin marks.");
            return GetFlt();
        }

        public bool IsZombieEmbed()
        {
            if (GetCh() != FIELD_END_MARK)
                throw new NotSupportedException(
                        "This field Is only defined for end marks.");
            return IsFZombieEmbed();
        }

        public bool IsResultDirty()
        {
            if (GetCh() != FIELD_END_MARK)
                throw new NotSupportedException(
                        "This field Is only defined for end marks.");
            return IsFResultDirty();
        }

        public bool IsResultEdited()
        {
            if (GetCh() != FIELD_END_MARK)
                throw new NotSupportedException(
                        "This field Is only defined for end marks.");
            return IsFResultEdited();
        }

        public bool IsLocked()
        {
            if (GetCh() != FIELD_END_MARK)
                throw new NotSupportedException(
                        "This field Is only defined for end marks.");
            return IsFLocked();
        }

        public bool IsPrivateResult()
        {
            if (GetCh() != FIELD_END_MARK)
                throw new NotSupportedException(
                        "This field Is only defined for end marks.");
            return IsFPrivateResult();
        }

        public bool IsNested()
        {
            if (GetCh() != FIELD_END_MARK)
                throw new NotSupportedException(
                        "This field Is only defined for end marks.");
            return IsFNested();
        }

        public bool IsHasSep()
        {
            if (GetCh() != FIELD_END_MARK)
                throw new NotSupportedException(
                        "This field Is only defined for end marks.");
            return IsFHasSep();
        }
    }
}

