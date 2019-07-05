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


namespace NPOI.HWPF.UserModel
{

    using System;
    using NPOI.HWPF.Model;
    public class Section : Range
    {

        private SectionProperties _props;

        public Section(SEPX sepx, Range parent)
            : base(Math.Max(parent._start, sepx.Start), Math.Min(parent._end, sepx.End), parent)
        {

            _props = sepx.GetSectionProperties();
        }

        public override int Type
        {
            get
            {
                return TYPE_SECTION;
            }
        }

        public int NumColumns
        {
            get
            {
                return _props.GetCcolM1() + 1;
            }
        }

        public override Object Clone()
        {
            Section s = (Section)base.Clone();
            s._props = (SectionProperties)_props.Clone();
            return s;
        }

        /**
         * @return distance to be maintained between columns, in twips. Used when
         *         {@link #isColumnsEvenlySpaced()} == true
         */
        public int DistanceBetweenColumns
        {
            get
            {
                return _props.GetDxaColumns();
            }
        }

        public int MarginBottom
        {
            get
            {
                return _props.GetDyaBottom();
            }
        }

        public int MarginLeft
        {
            get
            {
                return _props.GetDxaLeft();
            }
        }

        public int MarginRight
        {
            get
            { 
                return _props.GetDxaRight();
            }
        }

        public int MarginTop
        {
            get
            {
                return _props.GetDyaTop();
            }
        }

        /**
         * @return page height (in twips) in current section. Default value is 15840
         *         twips
         */
        public int PageHeight
        {
            get 
            {
                return _props.GetYaPage();
            }
        }

        /**
         * @return page width (in twips) in current section. Default value is 12240
         *         twips
         */
        public int PageWidth
        {
            get
            { 
                return _props.GetXaPage();
            }
        }

        public bool IsColumnsEvenlySpaced
        {
            get 
            {
                return _props.GetFEvenlySpaced();
            }
        }

        public override String ToString()
        {
            return "Section [" + StartOffset + "; " + EndOffset + ")";
        }
    }
}

