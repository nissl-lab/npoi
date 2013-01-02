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
using System.Text;
using System.Collections.Generic;

using NUnit.Framework;
using NPOI.Util;

namespace TestCases.Util
{
    /// <summary>
    /// Summary description for TestPOILoggerFactory
    /// </summary>
    [TestFixture]
    public class TestPOILogFactory
    {
        public TestPOILogFactory()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        [Test]
        public void TestLog()
        {
            //NKB Testing only that logging classes use gives no exception
            //    Since logging can be disabled, no checking of logging
            //    output is done.

            POILogger l1 = POILogFactory.GetLogger( "org.apache.poi.hssf.test" );
            POILogger l2 = POILogFactory.GetLogger( "org.apache.poi.hdf.test" );

            l1.Log( POILogger.FATAL, "testing cat org.apache.poi.hssf.*:FATAL" );
            l1.Log( POILogger.ERROR, "testing cat org.apache.poi.hssf.*:ERROR" );
            l1.Log( POILogger.WARN, "testing cat org.apache.poi.hssf.*:WARN" );
            l1.Log( POILogger.INFO, "testing cat org.apache.poi.hssf.*:INFO" );
            l1.Log( POILogger.DEBUG, "testing cat org.apache.poi.hssf.*:DEBUG" );

            l2.Log( POILogger.FATAL, "testing cat org.apache.poi.hdf.*:FATAL" );
            l2.Log( POILogger.ERROR, "testing cat org.apache.poi.hdf.*:ERROR" );
            l2.Log( POILogger.WARN, "testing cat org.apache.poi.hdf.*:WARN" );
            l2.Log( POILogger.INFO, "testing cat org.apache.poi.hdf.*:INFO" );
            l2.Log( POILogger.DEBUG, "testing cat org.apache.poi.hdf.*:DEBUG" );

        }
    }
}
