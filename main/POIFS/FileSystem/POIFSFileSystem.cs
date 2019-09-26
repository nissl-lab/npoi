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

using System;
using System.Collections;
using System.IO;

using NPOI.POIFS.Properties;
using NPOI.POIFS.Dev;
using NPOI.POIFS.Storage;
using NPOI.POIFS.EventFileSystem;
using NPOI.POIFS.Common;
using NPOI.Util;


namespace NPOI.POIFS.FileSystem
{
    /// <summary>
    /// This is the main class of the POIFS system; it manages the entire
    /// life cycle of the filesystem.
    /// @author Marc Johnson (mjohnson at apache dot org) 
    /// </summary>
    [Serializable]
    public class POIFSFileSystem :
        OPOIFSFileSystem // TODO Temporary workaround during #56791
        , POIFSViewable
    {

        /// <summary>
        /// Convenience method for clients that want to avoid the auto-Close behaviour of the constructor.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <example>
        /// A convenience method (
        /// CreateNonClosingInputStream()) has been provided for this purpose:
        /// StreamwrappedStream = POIFSFileSystem.CreateNonClosingInputStream(is);
        /// HSSFWorkbook wb = new HSSFWorkbook(wrappedStream);
        /// is.reset();
        /// doSomethingElse(is);
        /// </example>
        /// <returns></returns>
        public static Stream CreateNonClosingInputStream(Stream stream) {
            return new CloseIgnoringInputStream(stream);
        }


        /// <summary>
        /// Initializes a new instance of the <see cref="POIFSFileSystem"/> class.  intended for writing
        /// </summary>
        public POIFSFileSystem():base()
        {
            
        }

        /// <summary>
        /// Create a POIFSFileSystem from an Stream. Normally the stream is Read until
        /// EOF.  The stream is always Closed.  In the unlikely case that the caller has such a stream and
        /// needs to use it after this constructor completes, a work around is to wrap the
        /// stream in order to trap the Close() call.  
        /// </summary>
        /// <param name="stream">the Streamfrom which to Read the data</param>
        public POIFSFileSystem(Stream stream)
            : base(stream)
        {
            
        }

        /// <summary>
        /// Checks that the supplied Stream(which MUST
        /// support mark and reset, or be a PushbackInputStream)
        /// has a POIFS (OLE2) header at the start of it.
        /// If your Streamdoes not support mark / reset,
        /// then wrap it in a PushBackInputStream, then be
        /// sure to always use that, and not the original!
        /// </summary>
        /// <param name="inp">An Streamwhich supports either mark/reset, or is a PushbackStream</param>
        /// <returns>
        /// 	<c>true</c> if [has POIFS header] [the specified inp]; otherwise, <c>false</c>.
        /// </returns>
        public static new bool HasPOIFSHeader(Stream inp){

            return OPOIFSFileSystem.HasPOIFSHeader(inp);
        }
        /**
     * Checks if the supplied first 8 bytes of a stream / file
     *  has a POIFS (OLE2) header.
     */
        public static new bool HasPOIFSHeader(byte[] header8Bytes)
        {
            return OPOIFSFileSystem.HasPOIFSHeader(header8Bytes);
        }

    }
}
