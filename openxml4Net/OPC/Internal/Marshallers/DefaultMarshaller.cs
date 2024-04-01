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

namespace NPOI.OpenXml4Net.OPC.Internal.Marshallers
{
    /// <summary>
    /// Default marshaller that specified that the part is responsible to marshall its content.
    /// </summary>
    /// @see PartMarshaller
    /// <remarks>
    /// @author Julien Chable
    /// @version 1.0
    /// </remarks>

    public class DefaultMarshaller : PartMarshaller
    {

        /// <summary>
        /// Save part in the output stream by using the save() method of the part.
        /// </summary>
        /// <exception cref="OpenXml4NetException">OpenXml4NetException
        /// If any error occur.
        /// </exception>
        public bool Marshall(PackagePart part, Stream out1)
        {
            return part.Save(out1);
        }
    }
}
