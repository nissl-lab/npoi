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
using System.Collections.Generic;
using System.Text;

namespace NPOI.Util.Collections
{
    using System;
    using System.Collections;
    
    /// <summary>
    /// A List of objects that are indexed AND keyed by an int; also allows for getting
    /// the index of a value in the list
    /// I am happy is someone wants to re-implement this without using the
    /// internal list and hashmap. If so could you please make sure that
    /// you can add elements half way into the list and have the value-key mappings
    /// update
    /// @author Jason Height
    /// </summary>
    public class IntMapper
    {
        private ArrayList elements;
        private Hashtable valueKeyMap;

        private const int _default_size = 10;

        /// <summary>
        /// create an IntMapper of default size
        /// </summary>
        public IntMapper():this(_default_size)
        {
            
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="IntMapper"/> class.
        /// </summary>
        /// <param name="initialCapacity">The initial capacity.</param>
        public IntMapper(int initialCapacity)
        {
            elements = new ArrayList(initialCapacity);
            valueKeyMap = new Hashtable(initialCapacity);
        }

        /// <summary>
        /// Appends the specified element to the end of this list
        /// @param value element to be appended to this list.
        /// @return true (as per the general contract of the Collection.add
        /// method).
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns></returns>
        public bool Add(Object value)
        {
          int index = elements.Count;
          elements.Add(value);
          valueKeyMap[value]= index;
          return true;
        }

        /// <summary>
        /// Gets the size.
        /// </summary>
        /// <value>The size.</value>
        public int Size {
            get{
                  return elements.Count;
            }
        }

        /// <summary>
        /// Gets the <see cref="System.Object"/> at the specified index.
        /// </summary>
        /// <value></value>
        public virtual object this[int index]
        {
            get { return elements[index]; }
        }

        /// <summary>
        /// Gets the index.
        /// </summary>
        /// <param name="o">The o.</param>
        /// <returns></returns>
        public int GetIndex(Object o){
          object tmp = valueKeyMap[o];
          if (tmp == null)
          {
              return -1;
          }
          return (int)tmp;
        }

        /// <summary>
        /// Gets the enumerator.
        /// </summary>
        /// <returns></returns>
        public IEnumerator GetEnumerator() {
          return elements.GetEnumerator();
        }
    }
}
