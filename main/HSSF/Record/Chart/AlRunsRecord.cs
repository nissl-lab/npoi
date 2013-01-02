/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

/*
 * HSSF Chart Title Format Record Type
 */
namespace NPOI.HSSF.Record.Chart
{

    using System;
    using System.Collections;
    using System.Text;
    using NPOI.Util;


    /*
     * Describes the formatting runs associated with a chart title.
     */
    //
    /// <summary>
    /// The AlRuns record specifies Rich Text Formatting within chart 
    /// titles (section 2.2.3.3), trendline (section 2.2.3.12), and 
    /// data labels (section 2.2.3.11).
    /// </summary>
    public class AlRunsRecord : StandardRecord
    {
        public const short sid = 0x1050;

        private int m_recs;

        private class CTFormat
        {
            private short m_offset;
            private short m_fontIndex;

            public CTFormat(short offset, short fontIdx)
            {
                m_offset = offset;
                m_fontIndex = fontIdx;
            }

            public short Offset
            {
                get
                {
                    return m_offset;
                }
                set
                {
                    m_offset = value;
                }
            }
            public short FontIndex
            {
                get { return m_fontIndex; }
            }
            public void Serialize(ILittleEndianOutput out1)
            {
                out1.WriteShort(m_offset);
                out1.WriteShort(m_fontIndex);
            }
        }

        private ArrayList m_formats;

        public AlRunsRecord()
            : base()
        {

        }

        public AlRunsRecord(RecordInputStream in1)
        {
            m_recs = in1.ReadUShort();
            int idx;
            CTFormat ctf;
            if (m_formats == null)
            {
                m_formats = new ArrayList(m_recs);
            }
            for (idx = 0; idx < m_recs; idx++)
            {
                ctf = new CTFormat(in1.ReadShort(), in1.ReadShort());
                m_formats.Add(ctf);
            }
        }
        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(m_formats.Count);
            for (int i = 0; i < m_formats.Count; i++)
            {
                ((CTFormat)m_formats[i]).Serialize(out1);
            }
        }

        protected override int DataSize
        {
            get { return 2 + (4 * m_formats.Count); }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public int GetFormatCount()
        {
            return m_formats.Count;
        }

        public void ModifyFormatRun(short oldPos, short newLen)
        {
            short shift = (short)0;
            for (int idx = 0; idx < m_formats.Count; idx++)
            {
                CTFormat ctf = (CTFormat)m_formats[idx];
                if (shift != 0)
                {
                    ctf.Offset = ((short)(ctf.Offset + shift));
                }
                else if ((oldPos == ctf.Offset) && (idx < (m_formats.Count - 1)))
                {
                    CTFormat nextCTF = (CTFormat)m_formats[idx + 1];
                    shift = (short)(newLen - (nextCTF.Offset - ctf.Offset));
                }
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[ALRUNS]\n");
            buffer.Append("    .format_runs       = ").Append(m_recs)
                .Append("\n");
            int idx;
            CTFormat ctf;
            for (idx = 0; idx < m_formats.Count; idx++)
            {
                ctf = (CTFormat)m_formats[idx];
                buffer.Append("       .char_offset= ").Append(ctf.Offset);
                buffer.Append(",.fontidx= ").Append(ctf.FontIndex);
                buffer.Append("\n");
            }
            buffer.Append("[/ALRUNS]\n");
            return buffer.ToString();
        }
    }
}