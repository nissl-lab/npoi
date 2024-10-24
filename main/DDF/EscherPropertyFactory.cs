
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

            for (int i = 0; i < numProperties; i++)
            {
                short propId;
                int propData;
                propId = LittleEndian.GetShort(data, pos);
                propData = LittleEndian.GetInt(data, pos + 2);
                short propNumber = (short)(propId & (short)0x3FFF);
                bool isComplex = (propId & unchecked((short)0x8000)) != 0;

                byte propertyType = EscherProperties.GetPropertyType((short)propNumber);
                EscherProperty ep;
                switch (propertyType)
                {
                    case EscherPropertyMetaData.TYPE_BOOL:
                        ep = new EscherBoolProperty(propId, propData);
                        break;
                    case EscherPropertyMetaData.TYPE_RGB:
                        ep = new EscherRGBProperty(propId, propData);
                        break;
                    case EscherPropertyMetaData.TYPE_SHAPEPATH:
                        ep = new EscherShapePathProperty(propId, propData);
                        break;
                    default:
                        if (!isComplex)
                        {
                            ep = new EscherSimpleProperty(propId, propData);
                        }
                        else if (propertyType == EscherPropertyMetaData.TYPE_ARRAY)
                        {
                            ep = new EscherArrayProperty(propId, new byte[propData]);
                        }
                        else
                        {
                            ep = new EscherComplexProperty(propId, new byte[propData]);
                        }
                        break;
                }
                results.Add(ep);
                pos += 6;
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
                        int leftover = data.Length - pos;
                        if (leftover < complexData.Length)
                        {
                            throw new InvalidOperationException("Could not read complex escher property, lenght was " + complexData.Length + ", but had only " +
                                    leftover + " bytes left");
                        }
                        Array.Copy(data, pos, complexData, 0, complexData.Length);
                        pos += complexData.Length;
                    }
                }
            }

            return results;
        }


    }
}