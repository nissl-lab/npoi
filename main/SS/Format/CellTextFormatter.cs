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
namespace NPOI.SS.Format
{
    using System;
    using System.Text.RegularExpressions;
    using System.Text;


    /**
     * This class : printing out text.
     *
     * @author Ken Arnold, Industrious Media LLC
     */
    public class CellTextFormatter : CellFormatter
    {
        private int[] textPos;
        private String desc;

        internal static CellFormatter SIMPLE_TEXT = new CellTextFormatter("@");
        private class PartHandler : CellFormatPart.IPartHandler
        {
            private int numplace;
            public int NumPlace
            {
                get{return numplace;}
            }
            public PartHandler(int numPlace)
            {
                this.numplace = numPlace;
            }
            public String HandlePart(Match m, String part,
                                CellFormatType type, StringBuilder desc)
            {
                if (part.Equals("@"))
                {
                    numplace++;
                    return "\u0000";
                }
                return null;
            }
        }
        public CellTextFormatter(String format)
            : base(format)
        {
            ;

            int[] numPlaces = new int[1];
            PartHandler handler = new PartHandler(numPlaces[0]);
            desc = CellFormatPart.ParseFormat(format, CellFormatType.TEXT, handler).ToString();

            // Remember the "@" positions in last-to-first order (to make insertion easier)
            textPos = new int[handler.NumPlace];
            int pos = desc.Length - 1;
            for (int i = 0; i < textPos.Length; i++)
            {
                textPos[i] = desc.LastIndexOf("\u0000", pos);
                pos = textPos[i] - 1;
            }
        }

        /** {@inheritDoc} */
        public override void FormatValue(StringBuilder toAppendTo, Object obj)
        {
            int start = toAppendTo.Length;
            String text = obj.ToString();
            if (obj is Boolean)
            {
                text = text.ToUpper();
            }
            toAppendTo.Append(desc);
            for (int i = 0; i < textPos.Length; i++)
            {
                int pos = start + textPos[i];
                //toAppendTo.Replace(pos, pos + 1, text);
                toAppendTo.Remove(pos, 1).Insert(pos, text);
            }
        }

        /**
         * {@inheritDoc}
         * <p/>
         * For text, this is just printing the text.
         */
        public override void SimpleValue(StringBuilder toAppendTo, Object value)
        {
            SIMPLE_TEXT.FormatValue(toAppendTo, value);
        }
    }
}