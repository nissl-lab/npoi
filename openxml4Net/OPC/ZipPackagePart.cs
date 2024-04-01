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
using ICSharpCode.SharpZipLib.Zip;

namespace NPOI.OpenXml4Net.OPC
{
    /// <summary>
    /// Zip implementation of a PackagePart.
    /// </summary>
    /// @see PackagePart
    /// <remarks>
    /// @author Julien Chable
    /// @version 1.0
    /// </remarks>

    public class ZipPackagePart : PackagePart
    {

        /// <summary>
        /// The zip entry corresponding to this part.
        /// </summary>
        private ZipEntry zipEntry;

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="container">
        /// The container package.
        /// </param>
        /// <param name="partName">
        /// Part name.
        /// </param>
        /// <param name="contentType">
        /// Content type.
        /// </param>
        /// <exception cref="InvalidFormatException">
        /// Throws if the content of this part invalid.
        /// </exception>
        public ZipPackagePart(OPCPackage container, PackagePartName partName,
                String contentType) : base(container, partName, contentType)
        {

        }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="container">
        /// The container package.
        /// </param>
        /// <param name="zipEntry">
        /// The zip entry corresponding to this part.
        /// </param>
        /// <param name="partName">
        /// The part name.
        /// </param>
        /// <param name="contentType">
        /// Content type.
        /// </param>
        /// <exception cref="InvalidFormatException">
        /// Throws if the content of this part is invalid.
        /// </exception>
        public ZipPackagePart(OPCPackage container, ZipEntry zipEntry,
                PackagePartName partName, String contentType) : base(container, partName, contentType)
        {

            this.zipEntry = zipEntry;
        }

        /// <summary>
        /// Get the zip entry of this part.
        /// </summary>
        /// <returns>The zip entry in the zip structure coresponding to this part.</returns>
        public ZipEntry ZipArchive
        {
            get
            {
                return zipEntry;
            }
        }

        /// <summary>
        /// Implementation of the getInputStream() which return the inputStream of
        /// this part zip entry.
        /// </summary>
        /// <returns>Input stream of this part zip entry.</returns>

        protected override Stream GetInputStreamImpl()
        {
            // We use the getInputStream() method from java.util.zip.ZipFile
            // class which return an InputStream to this part zip entry.
            return ((ZipPackage) _container).ZipArchive
                    .GetInputStream(zipEntry);
        }

        protected override Stream GetOutputStreamImpl()
        {
            return null;
        }

        public override long Size
        {
            get
            {
                return zipEntry.Size;
            }
        }

        public override bool Save(Stream os)
        {
            return new ZipPartMarshaller().Marshall(this, os);
        }


        public override bool Load(Stream ios)
        {
            throw new InvalidOperationException("Method not implemented !");
        }


        public override void Close()
        {
            throw new InvalidOperationException("Method not implemented !");
        }


        public override void Flush()
        {
            throw new InvalidOperationException("Method not implemented !");
        }
    }
}
