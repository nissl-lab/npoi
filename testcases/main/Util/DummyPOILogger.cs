
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License Is distributed on an "AS Is" BASIS,
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
using NPOI.Util;

namespace TestCases.Util
{
    /// <summary>
    /// Summary description for DummyPOILogger
    /// </summary>
    public class DummyPOILogger:POILogger {
	    public ArrayList logged = new ArrayList(); 

	    public void Reset() {
		    logged.Clear(); // = new ArrayList();
	    }

        public override bool Check(int level)
        {
		    return true;
	    }

	    public override void Initialize(String cat) {}

        public override void Log(int level, Object obj1)
        {
		    logged.Add(level + " - " + obj1);
	    }

        public override void Log(int level, Object obj1, Exception exception)
        {
		    logged.Add(level + " - " + obj1 + " - " + exception);
	    }
    }
}
