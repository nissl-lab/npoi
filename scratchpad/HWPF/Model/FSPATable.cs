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

namespace NPOI.HWPF.Model
{
    using System;
    using System.Collections;
    using System.Text;
    using System.Collections.Generic;

    /**
     * This class holds all the FSPA (File Shape Address) structures.
     *
     * @author Squeeself
     */
    public class FSPATable
    {
        private ArrayList _shapes = new ArrayList();
        private Hashtable _shapeIndexesByPropertyStart = new Hashtable();
        private List<TextPiece> _text;

        public FSPATable(byte[] tableStream, int fcPlcspa, int lcbPlcspa, List<TextPiece> tpt)
        {
            _text = tpt;
            // Will be 0 if no drawing objects in document
            if (fcPlcspa == 0)
                return;

            PlexOfCps plex = new PlexOfCps(tableStream, fcPlcspa, lcbPlcspa, FSPA.FSPA_SIZE);
            for (int i = 0; i < plex.Length; i++)
            {
                GenericPropertyNode property = plex.GetProperty(i);
                FSPA fspa = new FSPA(property.Bytes, 0);

                _shapes.Add(fspa);
                _shapeIndexesByPropertyStart.Add(property.Start, i);
            }
        }

        public FSPA GetFspaFromCp(int cp)
        {
            if (!_shapeIndexesByPropertyStart.Contains(cp))
            {
                return null;
            }
            return (FSPA)_shapes[(int)_shapeIndexesByPropertyStart[cp]];
        }

        public FSPA[] GetShapes()
        {
            FSPA[] result = new FSPA[_shapes.Count];
            result = (FSPA[])_shapes.ToArray();
            return result;
        }

        public override String ToString()
        {
            StringBuilder buf = new StringBuilder();
            buf.Append("[FPSA PLC size=").Append(_shapes.Count).Append("]\n");
            for (IEnumerator it = _shapeIndexesByPropertyStart.Keys.GetEnumerator(); it.MoveNext(); )
            {
                int i = (int)it.Current;
                FSPA fspa = (FSPA)_shapes[(int)_shapeIndexesByPropertyStart[i]];
                buf.Append("  [FC: ").Append(i.ToString()).Append("] ");
                buf.Append(fspa.ToString());
                buf.Append("\n");
            }
            buf.Append("[/FSPA PLC]");
            return buf.ToString();
        }
    }
}

