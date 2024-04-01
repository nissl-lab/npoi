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
using NPOI.OpenXml4Net.OPC.Internal.Unmarshallers;

namespace NPOI.OpenXml4Net.OPC.Internal
{
    /// <summary>
    /// Object implemented this interface are considered as part unmarshaller. A part
    /// unmarshaller is responsible to unmarshall a part in order to load it from a
    /// package.
    /// </summary>
    /// <remarks>
    /// @author Julien Chable
    /// @version 0.1
    /// </remarks>

    public interface PartUnmarshaller
    {

        /// <summary>
        /// Save the content of the package in the stream
        /// </summary>
        /// <param name="in">
        /// The input stream from which the part will be unmarshall.
        /// </param>
        /// <returns>The part freshly unmarshall from the input stream.</returns>
        /// <exception cref="OpenXml4NetException">OpenXml4NetException
        /// Throws only if any other exceptions are thrown by inner
        /// methods.
        /// </exception>
        PackagePart Unmarshall(UnmarshallContext context, Stream in1);
    }
}
