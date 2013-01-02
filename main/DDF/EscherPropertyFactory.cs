
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

namespace NPOI.DDF
{
    using System;
    using System.Collections;
    using NPOI.Util;
    using System.Collections.Generic;

    /// <summary>
    /// Generates a property given a reference into the byte array storing that property.
    /// @author Glen Stampoultzis
    /// </summary>
    public class EscherPropertyFactory
    {
        /// <summary>
        /// Create new properties from a byte array.
        /// </summary>
        /// <param name="data">The byte array containing the property</param>
        /// <param name="offset">The starting offset into the byte array</param>
        /// <param name="numProperties">The new properties</param>
        /// <returns></returns>        
        public List<EscherProperty> CreateProperties(byte[] data, int offset, short numProperties)
        {
            List<EscherProperty> results = new List<EscherProperty>();

            int pos = offset;

            //        while ( bytesRemaining >= 6 )
            for (int i = 0; i < numProperties; i++)
            {
                short propId;
                int propData;
                propId = LittleEndian.GetShort(data, pos);
                propData = LittleEndian.GetInt(data, pos + 2);
                short propNumber = (short)(propId & (short)0x3FFF);
                bool isComplex = (propId & unchecked((short)0x8000)) != 0;
                bool isBlipId = (propId & (short)0x4000) != 0;

                byte propertyType = EscherProperties.GetPropertyType((short)propNumber);
                if (propertyType == EscherPropertyMetaData.TYPE_bool)
                    results.Add(new EscherBoolProperty(propId, propData));
                else if (propertyType == EscherPropertyMetaData.TYPE_RGB)
                    results.Add(new EscherRGBProperty(propId, propData));
                else if (propertyType == EscherPropertyMetaData.TYPE_SHAPEPATH)
                    results.Add(new EscherShapePathProperty(propId, propData));
                else
                {
                    if (!isComplex)
                        results.Add(new EscherSimpleProperty(propId, propData));
                    else
                    {
                        if (propertyType == EscherPropertyMetaData.TYPE_ARRAY)
                            results.Add(new EscherArrayProperty(propId, new byte[propData]));
                        else
                            results.Add(new EscherComplexProperty(propId, new byte[propData]));

                    }
                }
                pos += 6;
                //            bytesRemaining -= 6 + complexBytes;
            }

            // Get complex data
            for (IEnumerator iterator = results.GetEnumerator(); iterator.MoveNext(); )
            {
                EscherProperty p = (EscherProperty)iterator.Current;
                if (p is EscherComplexProperty)
                {
                    if (p is EscherArrayProperty)
                    {
                        pos += ((EscherArrayProperty)p).SetArrayData(data, pos);
                    }
                    else
                    {
                        byte[] complexData = ((EscherComplexProperty)p).ComplexData;
                        Array.Copy(data, pos, complexData, 0, complexData.Length);
                        pos += complexData.Length;
                    }
                }
            }

            return results;
        }


    }
}