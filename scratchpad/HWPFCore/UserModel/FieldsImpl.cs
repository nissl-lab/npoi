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
using NPOI.HWPF.Model;
using NPOI.HWPF.UserModel;
using System.Collections.Generic;
using System;
namespace NPOI.HWPF.UserModel
{

    /**
     * Default implementation of {@link Field}
     * 
     * @author Sergey Vladimirov (vlsergey {at} gmail {dot} com)
     */

    public class FieldsImpl : Fields
    {
        /**
         * This is port and adaptation of Arrays.binarySearch from Java 6 (Apache
         * Harmony).
         */
        private static int BinarySearch(List<PlexOfField> list,
                int startIndex, int endIndex, int requiredStartOffset)
        {
            CheckIndexForBinarySearch(list.Count, startIndex, endIndex);

            int low = startIndex, mid = -1, high = endIndex - 1, result = 0;
            while (low <= high)
            {
                mid = (low + high) >> 1;
                int midStart = list[mid].FcStart;

                if (midStart == requiredStartOffset)
                {
                    return mid;
                }
                else if (midStart < requiredStartOffset)
                {
                    low = mid + 1;
                }
                else
                {
                    high = mid - 1;
                }
            }
            if (mid < 0)
            {
                int insertPoint = endIndex;
                for (int index = startIndex; index < endIndex; index++)
                {
                    if (requiredStartOffset < list[index].FcStart)
                    {
                        insertPoint = index;
                    }
                }
                return -insertPoint - 1;
            }
            return -mid - (result >= 0 ? 1 : 2);
        }

        private static void CheckIndexForBinarySearch(int length, int start,
                int end)
        {
            if (start > end)
            {
                throw new ArgumentException();
            }
            if (length < end || 0 > start)
            {
                throw new IndexOutOfRangeException();
            }
        }

        private Dictionary<FieldsDocumentPart, Dictionary<int, FieldImpl>> _fieldsByOffset;

        private PlexOfFieldComparator comparator = new PlexOfFieldComparator();

        public FieldsImpl(FieldsTables fieldsTables)
        {
            _fieldsByOffset = new Dictionary<FieldsDocumentPart, Dictionary<int, FieldImpl>>(
                    );

            foreach (FieldsDocumentPart part in Enum.GetValues(typeof(FieldsDocumentPart)))
            {
                List<PlexOfField> plexOfCps = fieldsTables.GetFieldsPLCF(part);
                _fieldsByOffset.Add(part, ParseFieldStructure(plexOfCps));
            }
        }

        public List<Field> GetFields(FieldsDocumentPart part)
        {
            Dictionary<int, FieldImpl> map = _fieldsByOffset[part];
            if (map == null || map.Count == 0)
                return new List<Field>();

            List<Field> vList=new List<Field>();
            foreach(Field f in map.Values)
            {
                vList.Add(f);
            }
            return vList;
        }

        public Field GetFieldByStartOffset(FieldsDocumentPart documentPart,
                int offset)
        {
            Dictionary<int, FieldImpl> map = _fieldsByOffset[documentPart];
            if (map == null || map.Count == 0)
                return null;

            return map[offset];
        }

        private Dictionary<int, FieldImpl> ParseFieldStructure(
                List<PlexOfField> plexOfFields)
        {
            if (plexOfFields == null || plexOfFields.Count == 0)
                return new Dictionary<int, FieldImpl>();

            plexOfFields.Sort(comparator);
            List<FieldImpl> fields = new List<FieldImpl>(
                    plexOfFields.Count / 3 + 1);
            ParseFieldStructureImpl(plexOfFields, 0, plexOfFields.Count, fields);

            Dictionary<int, FieldImpl> result = new Dictionary<int, FieldImpl>(
                    fields.Count);
            foreach (FieldImpl field in fields)
            {
                result.Add(field.GetFieldStartOffset(), field);
            }
            return result;
        }

        private void ParseFieldStructureImpl(List<PlexOfField> plexOfFields,
                int startOffsetInclusive, int endOffsetExclusive,
                List<FieldImpl> result)
        {
            int next = startOffsetInclusive;
            while (next < endOffsetExclusive)
            {
                PlexOfField startPlexOfField = plexOfFields[next];
                if (startPlexOfField.Fld.GetBoundaryType() != FieldDescriptor.FIELD_BEGIN_MARK)
                {
                    /* Start mark seems to be missing */
                    next++;
                    continue;
                }

                /*
                 * we have start node. end offset points to next node, separator or
                 * end
                 */
                int nextNodePositionInList = BinarySearch(plexOfFields, next + 1,
                        endOffsetExclusive, startPlexOfField.FcEnd);
                if (nextNodePositionInList < 0)
                {
                    /*
                     * too bad, this start field mark doesn't have corresponding end
                     * field mark or separator field mark in fields table
                     */
                    next++;
                    continue;
                }
                PlexOfField nextPlexOfField = plexOfFields
                        [nextNodePositionInList];

                switch (nextPlexOfField.Fld.GetBoundaryType())
                {
                    case FieldDescriptor.FIELD_SEPARATOR_MARK:
                        {
                            PlexOfField separatorPlexOfField = nextPlexOfField;

                            int endNodePositionInList = BinarySearch(plexOfFields,
                                    nextNodePositionInList, endOffsetExclusive,
                                    separatorPlexOfField.FcEnd);
                            if (endNodePositionInList < 0)
                            {
                                /*
                                 * too bad, this separator field mark doesn't have
                                 * corresponding end field mark in fields table
                                 */
                                next++;
                                continue;
                            }
                            PlexOfField endPlexOfField = plexOfFields
                                    [endNodePositionInList];

                            if (endPlexOfField.Fld.GetBoundaryType() != FieldDescriptor.FIELD_END_MARK)
                            {
                                /* Not and ending mark */
                                next++;
                                continue;
                            }

                            FieldImpl field = new FieldImpl(startPlexOfField,
                                    separatorPlexOfField, endPlexOfField);
                            result.Add(field);

                            // Adding included fields
                            if (startPlexOfField.FcStart + 1 < separatorPlexOfField
                                    .FcStart - 1)
                            {
                                ParseFieldStructureImpl(plexOfFields, next + 1,
                                        nextNodePositionInList, result);
                            }
                            if (separatorPlexOfField.FcStart + 1 < endPlexOfField
                                    .FcStart - 1)
                            {
                                ParseFieldStructureImpl(plexOfFields,
                                        nextNodePositionInList + 1, endNodePositionInList,
                                        result);
                            }

                            next = endNodePositionInList + 1;

                            break;
                        }
                    case FieldDescriptor.FIELD_END_MARK:
                        {
                            // we have no separator
                            FieldImpl field = new FieldImpl(startPlexOfField, null,
                                    nextPlexOfField);
                            result.Add(field);

                            // Adding included fields
                            if (startPlexOfField.FcStart + 1 < nextPlexOfField
                                    .FcStart - 1)
                            {
                                ParseFieldStructureImpl(plexOfFields, next + 1,
                                        nextNodePositionInList, result);
                            }

                            next = nextNodePositionInList + 1;
                            break;
                        }
                    case FieldDescriptor.FIELD_BEGIN_MARK:
                    default:
                        {
                            /* something is wrong, ignoring this mark along with start mark */
                            next++;
                            continue;
                        }
                }
            }
        }

        private class PlexOfFieldComparator :
                IComparer<PlexOfField>
        {
            public int Compare(PlexOfField o1, PlexOfField o2)
            {
                int thisVal = o1.FcStart;
                int anotherVal = o2.FcStart;
                return thisVal < anotherVal ? -1 : thisVal == anotherVal ? 0 : 1;
            }
        }

    }
}