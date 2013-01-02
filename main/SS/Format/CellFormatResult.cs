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
    using System.Drawing;



    /**
     * This object Contains the result of Applying a cell format or cell format part
     * to a value.
     *
     * @author Ken Arnold, Industrious Media LLC
     * @see CellFormatPart#Apply(Object)
     * @see CellFormat#Apply(Object)
     */
    public class CellFormatResult
    {
        private bool _applies;
        private string _text;
        private Color _textcolor;
        /**
         * This is <tt>true</tt> if no condition was given that applied to the
         * value, or if the condition is satisfied.  If a condition is relevant, and
         * when applied the value fails the test, this is <tt>false</tt>.
         */
        public bool Applies
        {
            get { return _applies; }
            set { _applies = value; }
        }


        /** The resulting text.  This will never be <tt>null</tt>. */
        public String Text
        {
            get{return _text;}
            set{_text=value;}
        }

        /**
         * The color the format Sets, or <tt>null</tt> if the format Sets no color.
         * This will always be <tt>null</tt> if {@link #applies} is <tt>false</tt>.
         */
        public Color TextColor
        {
            get{return _textcolor;}
            set{_textcolor=value;}
        }

        /**
         * Creates a new format result object.
         *
         * @param applies   The value for {@link #applies}.
         * @param text      The value for {@link #text}.
         * @param textColor The value for {@link #textColor}.
         */
        public CellFormatResult(bool applies, String text, Color textColor)
        {
            this.Applies = applies;
            this.Text = text;
            this.TextColor = (applies ? textColor : Color.Empty);
        }
    }
}