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

using NPOI.Util;
using System.Collections.Generic;
using System.IO;
using System;
using NPOI.HWPF.Model.IO;
namespace NPOI.HWPF.Model
{


    /**
     * @author Ryan Ackley
     */
    public class ListTables
    {
        private static int LIST_DATA_SIZE = 28;
        private static int LIST_FORMAT_OVERRIDE_SIZE = 16;
        //private static POILogger log = POILogFactory.GetLogger(ListTables.class);

        Dictionary<int, ListData> _listMap = new Dictionary<int, ListData>();
        List<ListFormatOverride> _overrideList = new List<ListFormatOverride>();

        public ListTables()
        {

        }

        public ListTables(byte[] tableStream, int lstOffset, int lfoOffset)
        {
            // get the list data
            int length = LittleEndian.GetShort(tableStream, lstOffset);
            lstOffset += LittleEndianConsts.SHORT_SIZE;
            int levelOffset = lstOffset + (length * LIST_DATA_SIZE);

            for (int x = 0; x < length; x++)
            {
                ListData lst = new ListData(tableStream, lstOffset);
                _listMap.Add(lst.GetLsid(), lst);
                lstOffset += LIST_DATA_SIZE;

                int num = lst.numLevels();
                for (int y = 0; y < num; y++)
                {
                    ListLevel lvl = new ListLevel(tableStream, levelOffset);
                    lst.SetLevel(y, lvl);
                    levelOffset += lvl.GetSizeInBytes();
                }
            }

            // now get the list format overrides. The size is an int unlike the LST size
            length = LittleEndian.GetInt(tableStream, lfoOffset);
            lfoOffset += LittleEndianConsts.INT_SIZE;
            int lfolvlOffset = lfoOffset + (LIST_FORMAT_OVERRIDE_SIZE * length);
            for (int x = 0; x < length; x++)
            {
                ListFormatOverride lfo = new ListFormatOverride(tableStream, lfoOffset);
                lfoOffset += LIST_FORMAT_OVERRIDE_SIZE;
                int num = lfo.numOverrides();
                for (int y = 0; y < num; y++)
                {
                    while (tableStream[lfolvlOffset] == 255)
                    {
                        lfolvlOffset++;
                    }
                    ListFormatOverrideLevel lfolvl = new ListFormatOverrideLevel(tableStream, lfolvlOffset);
                    lfo.SetOverride(y, lfolvl);
                    lfolvlOffset += lfolvl.GetSizeInBytes();
                }
                _overrideList.Add(lfo);
            }
        }

        public int AddList(ListData lst, ListFormatOverride override1)
        {
            int lsid = lst.GetLsid();
            while (_listMap[lsid] != null)
            {
                lsid = lst.ResetListID();
                override1.SetLsid(lsid);
            }
            _listMap.Add(lsid, lst);
            _overrideList.Add(override1);
            return lsid;
        }

        public void WriteListDataTo(HWPFStream tableStream)
        {
            int listSize = _listMap.Count;

            // use this stream as a buffer for the levels since their size varies.
            MemoryStream levelBuf = new MemoryStream();

            byte[] shortHolder = new byte[2];
            LittleEndian.PutShort(shortHolder, (short)listSize);
            tableStream.Write(shortHolder);
            //TODO:: sort the keys
            foreach (int x in _listMap.Keys)
            {
                ListData lst = _listMap[x];
                tableStream.Write(lst.ToArray());
                ListLevel[] lvls = lst.GetLevels();
                for (int y = 0; y < lvls.Length; y++)
                {
                    byte[] bytes = lvls[y].ToArray();
                    levelBuf.Write(bytes, 0, bytes.Length);
                }
            }
            tableStream.Write(levelBuf.ToArray());
        }

        public void WriteListOverridesTo(HWPFStream tableStream)
        {

            // use this stream as a buffer for the levels since their size varies.
            MemoryStream levelBuf = new MemoryStream();

            int size = _overrideList.Count;

            byte[] intHolder = new byte[4];
            LittleEndian.PutInt(intHolder, size);
            tableStream.Write(intHolder);

            for (int x = 0; x < size; x++)
            {
                ListFormatOverride lfo = _overrideList[x];
                tableStream.Write(lfo.ToArray());
                ListFormatOverrideLevel[] lfolvls = lfo.GetLevelOverrides();
                for (int y = 0; y < lfolvls.Length; y++)
                {
                    byte[] bytes = lfolvls[y].ToArray();
                    levelBuf.Write(bytes, 0, bytes.Length);
                }
            }
            tableStream.Write(levelBuf.ToArray());

        }

        public ListFormatOverride GetOverride(int lfoIndex)
        {
            return _overrideList[lfoIndex - 1];
        }

        public int GetOverrideIndexFromListID(int lstid)
        {
            int returnVal = -1;
            int size = _overrideList.Count;
            for (int x = 0; x < size; x++)
            {
                ListFormatOverride next = _overrideList[x];
                if (next.GetLsid() == lstid)
                {
                    // 1-based index I think
                    returnVal = x + 1;
                    break;
                }
            }
            if (returnVal == -1)
            {
                throw new InvalidDataException("No list found with the specified ID");
            }
            return returnVal;
        }

        public ListLevel GetLevel(int listID, int level)
        {
            ListData lst = _listMap[listID];
            if (level < lst.numLevels())
            {
                ListLevel lvl = lst.GetLevels()[level];
                return lvl;
            }
            //log.log(POILogger.WARN, "Requested level " + level + " which was greater than the maximum defined (" + lst.numLevels() + ")");
            return null;
        }

        public ListData GetListData(int listID)
        {
            return _listMap[listID];
        }

        public override bool Equals(Object obj)
        {
            if (obj == null)
            {
                return false;
            }

            ListTables tables = (ListTables)obj;

            if (_listMap.Count == tables._listMap.Count)
            {
                foreach (int key in _listMap.Keys)
                {
                    ListData lst1 = _listMap[key];
                    ListData lst2 = tables._listMap[key];
                    if (!lst1.Equals(lst2))
                    {
                        return false;
                    }
                }
                int size = _overrideList.Count;
                if (size == tables._overrideList.Count)
                {
                    for (int x = 0; x < size; x++)
                    {
                        if (!_overrideList[x].Equals(tables._overrideList[x]))
                        {
                            return false;
                        }
                    }
                    return true;
                }
            }
            return false;
        }
    }
}