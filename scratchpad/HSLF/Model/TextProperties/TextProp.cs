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

using System;
namespace NPOI.HSLF.Model.TextProperties
{

    /** 
     * DefInition of a property of some text, or its paragraph. Defines 
     * how to find out if it's present (via the mask on the paragraph or 
     * character "Contains" header field), how long the value of it is, 
     * and how to Get and Set the value.
     * 
     * As the exact form of these (such as mask value, size of data
     *  block etc) is different for StyleTextProps and
     *  TxMasterTextProps, the defInitions of the standard
     *  TextProps is stored in the different record classes 
     */
    public class TextProp : ICloneable
    {
        protected int sizeOfDataBlock; // Number of bytes the data part uses
        protected String propName;
        protected int dataValue;
        protected int maskInHeader;

        /** 
         * Generate the defInition of a given type of text property.
         */
        public TextProp(int sizeOfDataBlock, int maskInHeader, String propName)
        {
            this.sizeOfDataBlock = sizeOfDataBlock;
            this.maskInHeader = maskInHeader;
            this.propName = propName;
            this.dataValue = 0;
        }

        /**
         * Name of the text property
         */
        public String GetName() { return propName; }

        /**
         * Size of the data section of the text property (2 or 4 bytes)
         */
        public int GetSize() { return sizeOfDataBlock; }

        /**
         * Mask in the paragraph or character "Contains" header field
         *  that indicates that this text property is present.
         */
        public int GetMask() { return maskInHeader; }
        /**
         * Get the mask that's used at write time. Only differs from
         *  the result of GetMask() for the mask based properties 
         */
        public virtual int GetWriteMask() { return GetMask(); }

        /**
         * Fetch the value of the text property (meaning is specific to
         *  each different kind of text property)
         */
        public int GetValue() { return dataValue; }

        /**
         * Set the value of the text property.
         */
        public virtual void SetValue(int val) { dataValue = val; }

        /**
         * Clone, eg when you want to actually make use of one of these.
         */
        public virtual Object Clone()
        {
            TextProp tp = new TextProp(this.sizeOfDataBlock, this.maskInHeader, this.propName);
            return tp;
        }
    }

}