/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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
using System.Collections.Generic;
using System.IO;
using NPOI.POIFS.Common;
using NPOI.POIFS.FileSystem;
using NPOI.POIFS.Storage;
using NPOI.Util;

namespace NPOI.POIFS.Properties
{
    public class NPropertyTable : PropertyTableBase
    {
        private POIFSBigBlockSize _bigBigBlockSize;

        public NPropertyTable(HeaderBlock headerBlock) : base(headerBlock)
        {
            _bigBigBlockSize = headerBlock.BigBlockSize;
        }

        public NPropertyTable(HeaderBlock headerBlock, NPOIFSFileSystem fileSystem)
            :base(headerBlock, 
                    BuildProperties( (new NPOIFSStream(fileSystem, headerBlock.PropertyStart)).GetEnumerator(),headerBlock.BigBlockSize)
            )
        {
            _bigBigBlockSize = headerBlock.BigBlockSize;
        }

        private static List<Property> BuildProperties(IEnumerator<ByteBuffer> dataSource, POIFSBigBlockSize bigBlockSize)
        {
            List<Property> properties = new List<Property>();

            while(dataSource.MoveNext())
            {
                ByteBuffer bb = dataSource.Current;

                // Turn it into an array
                byte[] data;
                if (bb.HasBuffer && bb.Offset == 0 &&
                    bb.Buffer.Length == bigBlockSize.GetBigBlockSize())
                {
                    data = bb.Buffer;
                }
                else
                {
                    data = new byte[bigBlockSize.GetBigBlockSize()];
                    int toRead = data.Length;
                    if (bb.Remaining() < bigBlockSize.GetBigBlockSize())
                    {
                        // Looks to be a truncated block
                        // This isn't allowed, but some third party created files
                        //  sometimes do this, and we can normally read anyway

                        toRead = bb.Remaining();
                    }
                    bb.Read(data, 0, toRead);
                }

                PropertyFactory.ConvertToProperties(data, properties);
            }
            return properties;

        }

        public override int CountBlocks
        {
            get
            {
                int size = _properties.Count * POIFSConstants.PROPERTY_SIZE;
                
                return (int)Math.Ceiling(1.0*size/_bigBigBlockSize.GetBigBlockSize());
            }
        }
        /**
     * Prepare to be written
     */
        public void PreWrite()
        {
            List<Property> pList = new List<Property>();
            // give each property its index
            int i = 0;
            foreach (Property p in _properties)
            {
                // only handle non-null properties 
                if (p == null) continue;
                p.Index = (i++);
                pList.Add(p);
            }

            // prepare each property for writing
            foreach (Property p in pList) p.PreWrite();
        }

        public void Write(NPOIFSStream stream)
        {
            Stream os = stream.GetOutputStream();
            try
            {
                //Leon ByteArrayOutputStream  -->MemoryStream
                MemoryStream ms = new MemoryStream();
                foreach (Property property in _properties)
                {
                    if (property != null)
                        property.WriteData(os);
                }

                os.Close();

                // Update the start position if needed
                if (StartBlock != stream.GetStartBlock())
                {
                    StartBlock = stream.GetStartBlock();
                }
            }
            catch (System.IO.IOException ex)
            {
                throw ex;
            }
        }
    }
}
