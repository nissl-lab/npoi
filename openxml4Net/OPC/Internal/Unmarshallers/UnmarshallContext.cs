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
using ICSharpCode.SharpZipLib.Zip;

namespace NPOI.OpenXml4Net.OPC.Internal.Unmarshallers
{
    /// <summary>
    /// Context needed for the unmarshall process of a part. This class is immutable.
    /// </summary>
    /// <remarks>
    /// @author Julien Chable
    /// @version 1.0
    /// </remarks>

    public class UnmarshallContext
    {

        private OPCPackage _package;

        private PackagePartName partName;

        private ZipEntry zipEntry;

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="targetPackage">Container.
        /// </param>
        /// <param name="partName">Name of the part to unmarshall.
        /// </param>
        public UnmarshallContext(OPCPackage targetPackage, PackagePartName partName)
        {
            this._package = targetPackage;
            this.partName = partName;
        }

        /// <summary>
        /// </summary>
        /// <returns>the container</returns>
        internal OPCPackage Package
        {
            get
            {
                return _package;
            }
            set 
            {
                this._package = value;
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>the partName</returns>
        internal PackagePartName PartName
        {
            get
            {
                return partName;
            }
            set
            {
                this.partName = value;
            }
        }
        /// <summary>
        /// </summary>
        /// <returns>the zipEntry</returns>
        internal ZipEntry ZipEntry
        {
            get
            {
                return zipEntry;
            }
            set
            {
                this.zipEntry = value;
            }
        }
    }
}
