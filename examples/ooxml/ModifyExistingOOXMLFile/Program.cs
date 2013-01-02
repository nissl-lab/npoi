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

/* ================================================================
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * NPOI HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/

using System;
using System.Collections.Generic;
using System.Text;

using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using NPOI.OpenXml4Net.OPC;
using System.Net.Mime;

namespace NPOI.Examples.ModifyExistingOOXMLFile
{
    class Program
    {
        static void Main(string[] args)
        {
            //create ooxml file in memory
            Package p = Package.Open("test.zip", PackageAccess.READ_WRITE);

            PackagePartName pn3 = new PackagePartName(new Uri("/c.xml", UriKind.Relative), true);
            if (!p.ContainPart(pn3))
                p.CreatePart(pn3, MediaTypeNames.Text.Xml);

            //save file 
            p.Save("test1.zip");

            //don't forget to close it
            p.Close();

        }

    }
}
