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
     * DefInition of a special kind of property of some text, or its 
     *  paragraph. For these properties, a flag in the "Contains" header 
     *  field tells you the data property family will exist. The value
     *  of the property is itself a mask, encoding several different
     *  (but related) properties
     */
    public class BitMaskTextProp : TextProp, ICloneable
    {
        private String[] subPropNames;
        private int[] subPropMasks;
        private bool[] subPropMatches;

        /** Fetch the list of the names of the sub properties */
        public String[] GetSubPropNames() { return subPropNames; }
        /** Fetch the list of if the sub properties match or not */
        public bool[] GetSubPropMatches() { return subPropMatches; }

        public BitMaskTextProp(int sizeOfDataBlock, int maskInHeader, String overallName, String[] subPropNames)
            : base(sizeOfDataBlock, maskInHeader, "bitmask")
        {

            this.subPropNames = subPropNames;
            this.propName = overallName;
            subPropMasks = new int[subPropNames.Length];
            subPropMatches = new bool[subPropNames.Length];

            // Initialise the masks list
            for (int i = 0; i < subPropMasks.Length; i++)
            {
                subPropMasks[i] = (1 << i);
            }
        }

        /**
         * As we're purely mask based, just Set flags for stuff
         *  that is Set
         */
        public override int GetWriteMask()
        {
            return dataValue;
        }

        /**
         * Set the value of the text property, and recompute the sub
         *  properties based on it
         */
        public override void SetValue(int val)
        {
            dataValue = val;

            // Figure out the values of the sub properties
            for (int i = 0; i < subPropMatches.Length; i++)
            {
                subPropMatches[i] = false;
                if ((dataValue & subPropMasks[i]) != 0)
                {
                    subPropMatches[i] = true;
                }
            }
        }

        /**
         * Fetch the true/false status of the subproperty with the given index
         */
        public bool GetSubValue(int idx)
        {
            return subPropMatches[idx];
        }

        /**
         * Set the true/false status of the subproperty with the given index
         */
        public void SetSubValue(bool value, int idx)
        {
            if (subPropMatches[idx] == value) { return; }
            if (value)
            {
                dataValue += subPropMasks[idx];
            }
            else
            {
                dataValue -= subPropMasks[idx];
            }
            subPropMatches[idx] = value;
        }

        public override Object Clone()
        {
            BitMaskTextProp newObj = (BitMaskTextProp)base.Clone();

            // Don't carry over matches, but keep everything 
            //  else as it was
            newObj.subPropMatches = new bool[subPropMatches.Length];

            return newObj;
        }
    }
}
