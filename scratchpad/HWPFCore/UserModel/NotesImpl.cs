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
using System.Collections.Generic;
using NPOI.HWPF.Model;
namespace NPOI.HWPF.UserModel
{

    /**
     * Default implementation of {@link Notes} interface
     * 
     * @author Sergey Vladimirov (vlsergey {at} gmail {doc} com)
     */
    public class NotesImpl : Notes
    {
        private Dictionary<int, int> anchorToIndexMap = null;

        private NotesTables notesTables;

        public NotesImpl(NotesTables notesTables)
        {
            this.notesTables = notesTables;
        }

        public int GetNoteAnchorPosition(int index)
        {
            return notesTables.GetDescriptor(index).Start;
        }

        public int GetNoteIndexByAnchorPosition(int anchorPosition)
        {
            UpdateAnchorToIndexMap();

            if(!anchorToIndexMap.ContainsKey(anchorPosition))
                return -1; 
            int index = anchorToIndexMap
                    [anchorPosition];               

            return index;
        }

        public int GetNotesCount()
        {
            return notesTables.GetDescriptorsCount();
        }

        public int GetNoteTextEndOffSet(int index)
        {
            return notesTables.GetTextPosition(index).End;
        }

        public int GetNoteTextStartOffSet(int index)
        {
            return notesTables.GetTextPosition(index).Start;
        }

        private void UpdateAnchorToIndexMap()
        {
            if (anchorToIndexMap != null)
                return;

            Dictionary<int, int> result = new Dictionary<int, int>();
            for (int n = 0; n < notesTables.GetDescriptorsCount(); n++)
            {
                int anchorPosition = notesTables.GetDescriptor(n).Start;
                result[anchorPosition] = n;
            }
            this.anchorToIndexMap = result;
        }
    }
}

