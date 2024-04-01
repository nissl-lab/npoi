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

namespace NPOI.OpenXml4Net.OPC
{
    /// <summary>
    /// A package part collection.
    /// </summary>
    /// <remarks>
    /// @author Julien Chable
    /// @version 0.1
    /// </remarks>

    public class PackagePartCollection : SortedList<PackagePartName, PackagePart>
    {
        /// <summary>
        /// Arraylist use to store this collection part names as string for rule
        /// M1.11 optimized checking.
        /// </summary>
        private List<String> registerPartNameStr = new List<String>();


        /// <summary>
        /// Check rule [M1.11]: a package implementer shall neither create nor
        /// recognize a part with a part name derived from another part name by
        /// Appending segments to it.
        /// </summary>
        /// <exception cref="InvalidOperationException">
        /// Throws if you try to add a part with a name derived from
        /// another part name.
        /// </exception>
        public PackagePart Put(PackagePartName partName, PackagePart part)
        {
            String[] segments = partName.URI.OriginalString.Split(
                    PackagingUriHelper.FORWARD_SLASH_CHAR);
            StringBuilder concatSeg = new StringBuilder();
            foreach (String seg in segments)
            {
                if (!seg.Equals(""))
                    concatSeg.Append(PackagingUriHelper.FORWARD_SLASH_CHAR);
                concatSeg.Append(seg);
                if (this.registerPartNameStr.Contains(concatSeg.ToString()))
                {
                    throw new InvalidOperationException(
                            "You can't add a part with a part name derived from another part ! [M1.11]");
                }
            }
            this.registerPartNameStr.Add(partName.Name);
            return base[partName] = part;
        }

        public new void Remove(PackagePartName key)
        {
            this.registerPartNameStr.Remove(((PackagePartName)key).Name);
            base.Remove(key);
        }
    }

}
