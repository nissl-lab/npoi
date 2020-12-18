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
        private Dictionary<int, GenericPropertyNode> _byStart
            = new Dictionary<int, GenericPropertyNode>();

        public FSPATable(byte[] tableStream, FileInformationBlock fib,
        FSPADocumentPart part)
        {
            int offset = fib.GetFSPAPlcfOffset(part);
            int length = fib.GetFSPAPlcfLength(part);

            PlexOfCps plex = new PlexOfCps(tableStream, offset, length,
                    FSPA.FSPA_SIZE);
            for (int i = 0; i < plex.Length; i++)
            {
                GenericPropertyNode property = plex.GetProperty(i);
                _byStart.Add(property.Start, property);
            }
        }


        [Obsolete]
        public FSPATable(byte[] tableStream, int fcPlcspa, int lcbPlcspa, List<TextPiece> tpt)
        {
            // Will be 0 if no drawing objects in document
            if (fcPlcspa == 0)
                return;

            PlexOfCps plex = new PlexOfCps(tableStream, fcPlcspa, lcbPlcspa, FSPA.FSPA_SIZE);
            for (int i = 0; i < plex.Length; i++)
            {
                GenericPropertyNode property = plex.GetProperty(i);
                _byStart.Add(property.Start, property);
            }
        }

        public FSPA GetFspaFromCp(int cp)
        {
            if (!_byStart.ContainsKey(cp))
            {
                return null;
            }

            return new FSPA(_byStart[cp].Bytes, 0);
        }

        public FSPA[] GetShapes()
        {
            List<FSPA> result = new List<FSPA>(_byStart.Count);
            foreach (GenericPropertyNode propertyNode in _byStart.Values)
            {
                result.Add(new FSPA(propertyNode.Bytes, 0));
            }
            return result.ToArray();
        }

        public override String ToString()
        {
            StringBuilder buf = new StringBuilder();
            buf.Append("[FPSA PLC size=").Append(_byStart.Count).Append("]\n");
            foreach (KeyValuePair<int, GenericPropertyNode> entry in _byStart
                     )
            {
                int i = entry.Key;
                buf.Append("  ").Append(i.ToString()).Append(" => \t");

                try
                {
                    FSPA fspa = GetFspaFromCp(i);
                    buf.Append(fspa.ToString());
                }
                catch (Exception exc)
                {
                    buf.Append(exc.Message);
                }
                buf.Append("\n");
            }
            buf.Append("[/FSPA PLC]");
            return buf.ToString();
        }
    }
}

