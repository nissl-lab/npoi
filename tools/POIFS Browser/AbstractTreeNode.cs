
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
 * POIFS Browser 
 * Author: NPOI Team
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * Huseyin Tufekcilerli     2008.11         1.0
 * Tony Qu                  2009.2.18       1.2 alpha
 * 
 * ==============================================================*/

using System;
using System.Windows.Forms;
using NPOI.POIFS.FileSystem;

namespace NPOI.Tools.POIFSBrowser
{
    internal abstract class AbstractTreeNode : TreeNode, IComparable<AbstractTreeNode>
    {
        public EntryNode EntryNode { get; private set; }

        public AbstractTreeNode(EntryNode entryNode)
            : base(entryNode.Name)
        {
            this.EntryNode = entryNode;
        }

        protected void ChangeImage(string imageKey)
        {
            this.ImageKey = this.SelectedImageKey = imageKey;
        }

        #region IComparable<AbstractTreeNode> Members

        public int CompareTo(AbstractTreeNode other)
        {
            bool b1 = this  is DirectoryTreeNode;
            bool b2 = other is DirectoryTreeNode;

            if (b1 && !b2)
            {
                return -1;
            }
            else if (!b1 && b2)
            {
                return 1;
            }
            else
            {
                return this.Text.CompareTo(other.Text);
            }
        }

        #endregion
    }
}