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

namespace TestCases.OpenXml4Net.OPC.Internal
{
    using NPOI.OpenXml4Net.OPC.Internal;
    using NPOI.OpenXml4Net.OPC.Internal.Marshallers;
    using NUnit.Framework;
    using System;
    using System.IO;

    [TestFixture]
    public class TestZipPackagePropertiesMarshaller
    {

        [Test]//(expected=ArgumentException.class)
        public void NonZipOutputStream()
        {
            PartMarshaller marshaller = new ZipPackagePropertiesMarshaller();
            Stream notAZipOutputStream = new MemoryStream(0);
            try
            {
                marshaller.Marshall(null, notAZipOutputStream);
            }
            catch (ArgumentException ex)
            {
                Assert.AreEqual("ZipOutputStream expected!", ex.Message);
            }
            
        }

    }

}