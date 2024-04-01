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
using ICSharpCode.SharpZipLib.Zip;

namespace NPOI.OpenXml4Net.OPC.Internal.Marshallers
{
    /// <summary>
    /// Package core properties marshaller specialized for zipped package.
    /// </summary>
    /// <remarks>
    /// @author Julien Chable
    /// </remarks>

    public class ZipPackagePropertiesMarshaller : PackagePropertiesMarshaller
    {
        public override bool Marshall(PackagePart part, Stream out1)
        {
            if(!(out1 is ZipOutputStream))
            {
                throw new ArgumentException("ZipOutputStream expected!");
            }
            ZipOutputStream zos = (ZipOutputStream) out1;

            // Saving the part in the zip file
            string name = ZipHelper
                .GetZipItemNameFromOPCName(part.PartName.URI.ToString());
            ZipEntry ctEntry = new ZipEntry(name);

            try
            {
                // Save in ZIP
                zos.PutNextEntry(ctEntry); // Add entry in ZIP

                base.Marshall(part, out1); // Marshall the properties inside a XML
                                           // Document
                StreamHelper.SaveXmlInStream(xmlDoc, out1);

                zos.CloseEntry();
            }
            catch(IOException e)
            {
                throw new OpenXml4NetException(e.Message, e);
            }
            catch
            {
                return false;
            }
            return true;
        }
    }

}
