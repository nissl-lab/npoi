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

/* ================================================================
 * About NPOI
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * Author's Blog: tonyqus.wordpress.com.cn (wp.tonyqus.cn)
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/

using System.Collections.Generic;

using NPOI.POIFS.Storage;
using NPOI.POIFS.Common;

namespace NPOI.POIFS.Properties
{
    public class PropertyFactory
    {
        private PropertyFactory()
        {
        }

        /// <summary>
        /// Convert raw data blocks to an array of Property's
        /// </summary>
        /// <param name="blocks">The blocks to be converted</param>
        /// <returns>the converted List of Property objects. May contain
        /// nulls, but will not be null</returns>
        public static List<Property> ConvertToProperties(ListManagedBlock [] blocks)
        {
            List<Property> properties = new List<Property>();

            for (int j = 0; j < blocks.Length; j++)
            {
                byte[] data = blocks[j].Data;
                ConvertToProperties(data, properties);
            }

            return properties;
        }

        public static void ConvertToProperties(byte[] data, List<Property> properties)
        {
            int property_count = data.Length / POIFSConstants.PROPERTY_SIZE;
                int    offset         = 0;

                for (int k = 0; k < property_count; k++)
                {
                switch (data[offset + PropertyConstants.PROPERTY_TYPE_OFFSET])
                    {
                    case PropertyConstants.DIRECTORY_TYPE:
                        properties.Add(new DirectoryProperty(properties.Count, data, offset));
                            break;
                    case PropertyConstants.DOCUMENT_TYPE:
                        properties.Add(new DocumentProperty(properties.Count, data, offset));
                            break;
                    case PropertyConstants.ROOT_TYPE:
                        properties.Add(new RootProperty(properties.Count, data, offset));
                            break;
                    default:
                            properties.Add(null);
                            break;
                    }
                    offset += POIFSConstants.PROPERTY_SIZE;
                }
        }
    }
}
