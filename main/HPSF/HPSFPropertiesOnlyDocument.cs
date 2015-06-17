/* ====================================================================
   Licensed To the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file To You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed To in writing, software
   distributed under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

using NPOI.POIFS.FileSystem;
using System.IO;
using System;
namespace NPOI.HPSF
{

    /**
     * A version of {@link POIDocument} which allows access to the
     *  HPSF Properties, but no other document contents.
     * Normally used when you want to read or alter the Document Properties,
     *  without affecting the rest of the file
     */
    public class HPSFPropertiesOnlyDocument : POIDocument
    {
        public HPSFPropertiesOnlyDocument(NPOIFSFileSystem fs)
            : base(fs.Root)
        {

        }
        public HPSFPropertiesOnlyDocument(POIFSFileSystem fs)
            : base(fs)
        {

        }

        public override void Write(Stream out1)
        {
            throw new InvalidOperationException("Unable to write, only for properties!");
        }
    }
}