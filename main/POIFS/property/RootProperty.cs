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

using NPOI.POIFS.Common;
using NPOI.POIFS.Storage;

namespace NPOI.POIFS.Properties
{
    public class RootProperty:DirectoryProperty
    {
        private const string NAME = "Root Entry";

        public RootProperty():base(NAME)
        {
            this.NodeColor=_NODE_BLACK;
            this.PropertyType=PropertyConstants.ROOT_TYPE;
            this.StartBlock = POIFSConstants.END_OF_CHAIN;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RootProperty"/> class.
        /// </summary>
        /// <param name="index">index number</param>
        /// <param name="array">byte data</param>
        /// <param name="offset">offset into byte data</param>
        public RootProperty(int index, byte [] array,
                               int offset): base(index, array, offset)
        {
           
        }

        /// <summary>
        /// Gets or sets the size of the document associated with this Property
        /// </summary>
        /// <value>the size of the document, in bytes</value>
        public override int Size
        {
            set{
                base.Size=SmallDocumentBlock.CalcSize(value);
            }
        }
    }
}
