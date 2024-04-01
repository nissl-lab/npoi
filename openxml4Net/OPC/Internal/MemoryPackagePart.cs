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
using System.Text;
using System.IO;
using NPOI.OpenXml4Net.OPC.Internal.Marshallers;
using NPOI.OpenXml4Net.Exceptions;
using NPOI.Util;

namespace NPOI.OpenXml4Net.OPC.Internal
{
    public class MemoryPackagePart : PackagePart
    {
        /// <summary>
        /// Storage for the part data.
        /// </summary>
        internal MemoryStream data;

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="pack">
        /// The owner package.
        /// </param>
        /// <param name="partName">
        /// The part name.
        /// </param>
        /// <param name="contentType">
        /// The content type.
        /// </param>
        /// <exception cref="InvalidFormatException">InvalidFormatException
        /// If the specified URI is not OPC compliant.
        /// </exception>
        public MemoryPackagePart(OPCPackage pack, PackagePartName partName,
                String contentType)
            : base(pack, partName, contentType)
        {

        }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="pack">
        /// The owner package.
        /// </param>
        /// <param name="partName">
        /// The part name.
        /// </param>
        /// <param name="contentType">
        /// The content type.
        /// </param>
        /// <param name="loadRelationships">loadRelationships
        /// Specify if the relationships will be loaded.
        /// </param>
        /// <exception cref="InvalidFormatException">InvalidFormatException
        /// If the specified URI is not OPC compliant.
        /// </exception>
        public MemoryPackagePart(OPCPackage pack, PackagePartName partName,
                String contentType, bool loadRelationships) :
            base(pack, partName, new ContentType(contentType), loadRelationships)
        {

        }

        protected override Stream GetInputStreamImpl()
        {
            // If this part has been created from scratch and/or the data buffer is
            // not
            // initialize, so we do it now.
            if (data == null)
            {
                return new MemoryStream();
            }
            MemoryStream newMs = new MemoryStream((int)data.Length);
            data.Position = 0;
            StreamHelper.CopyStream(data, newMs);
            newMs.Position = 0;
            return newMs;
        }

        protected override Stream GetOutputStreamImpl()
        {
            return new MemoryPackagePartOutputStream(this);
        }

        public override long Size
        {
            get
            {
                return data == null ? 0 : data.Length;
            }
        }
        public override void Clear()
        {
            data = null;
        }

        public override bool Save(Stream os)
        {
            return new ZipPartMarshaller().Marshall(this, os);
        }

        public override bool Load(Stream ios)
        {
            // Save it
            StreamHelper.CopyStream(ios, data);
            // All done
            return true;
        }

        public override void Close()
        {
            // Do nothing
        }

        public override void Flush()
        {
            // Do nothing
        }
    }
}
