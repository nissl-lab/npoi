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
using System.Xml;
using ICSharpCode.SharpZipLib.Zip;
using NPOI.Util;
namespace NPOI.OpenXml4Net.OPC.Internal
{
    /// <summary>
    /// Zip implementation of the ContentTypeManager.
    /// </summary>
    /// @see ContentTypeManager
    /// <remarks>
    /// @author Julien Chable
    /// @version 1.0
    /// </remarks>

    public class ZipContentTypeManager : ContentTypeManager
    {

        private static POILogger logger = POILogFactory.GetLogger(typeof(ZipContentTypeManager));
        /// <summary>
        /// Delegate constructor to the super constructor.
        /// </summary>
        /// <param name="in">
        /// The input stream to parse to fill internal content type
        /// collections.
        /// </param>
        /// <exception cref="Exceptions.InvalidFormatException">
        /// If the content types part content is not valid.
        /// </exception>
        public ZipContentTypeManager(Stream in1, OPCPackage pkg) : base(in1, pkg)
        {

        }


        public override bool SaveImpl(XmlDocument content, Stream out1)
        {
            ZipOutputStream zos = null;
            if(out1 is ZipOutputStream)
                zos = (ZipOutputStream) out1;
            else
                zos = new ZipOutputStream(out1);

            ZipEntry partEntry = new ZipEntry(CONTENT_TYPES_PART_NAME);
            try
            {
                // Referenced in ZIP
                zos.PutNextEntry(partEntry);
                // Saving data in the ZIP file

                StreamHelper.SaveXmlInStream(content, out1);
                Stream ins =  new MemoryStream();

                byte[] buff = new byte[ZipHelper.READ_WRITE_FILE_BUFFER_SIZE];
                while(true)
                {
                    int resultRead = ins.Read(buff, 0, ZipHelper.READ_WRITE_FILE_BUFFER_SIZE);
                    if(resultRead == 0)
                    {
                        // end of file reached
                        break;
                    }
                    else
                    {
                        zos.Write(buff, 0, resultRead);
                    }
                }
                zos.CloseEntry();
            }
            catch(IOException ioe)
            {
                logger.Log(POILogger.ERROR, "Cannot write: " + CONTENT_TYPES_PART_NAME
                    + " in Zip !", ioe);
                return false;
            }
            return true;
        }
    }

}
