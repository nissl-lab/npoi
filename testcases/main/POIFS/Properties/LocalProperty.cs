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
using NPOI.POIFS.Properties;

namespace TestCases.POIFS.Properties
{
    internal class LocalProperty : Property
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="LocalProperty"/> class.
        /// </summary>
        /// <param name="index">The index.</param>
        public LocalProperty(int index):base()
        {
            this.Name="foo" + index;
            this.Index=index;
        }
        /// <summary>
        /// Initializes a new instance of the <see cref="LocalProperty"/> class.
        /// </summary>
        /// <param name="name">name of the property</param>
        public LocalProperty(String name)
            : base()
        {
            this.Name = name;
        }

        /// <summary>
        /// do nothing
        /// </summary>
        public override void PreWrite()
        {
        }

        /// <summary>
        /// Gets a value indicating whether this instance is directory.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if a directory type Property; otherwise, <c>false</c>.
        /// </value>
        /// @return false
        public override bool IsDirectory
        {
            get { return false; }
        }
    }

}