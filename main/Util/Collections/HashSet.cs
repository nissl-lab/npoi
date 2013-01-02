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
using System.Collections;

namespace NPOI.Util.Collections
{
    /// <summary>
    /// This class comes from Java
    /// </summary>
	public class HashSet: ISet
	{
		private readonly Hashtable impl = new Hashtable();

        /// <summary>
        /// Initializes a new instance of the <see cref="HashSet"/> class.
        /// </summary>
		public HashSet()
		{
		}

        /// <summary>
        /// Initializes a new instance of the <see cref="HashSet"/> class.
        /// </summary>
        /// <param name="s">The s.</param>
		public HashSet(ISet s)
		{
			foreach (object o in s)
			{
				Add(o);
			}
		}

        /// <summary>
        /// Adds the specified o.
        /// </summary>
        /// <param name="o">The o.</param>
		public void Add(object o)
		{
			impl[o] = null;
		}

        /// <summary>
        /// Determines whether [contains] [the specified o].
        /// </summary>
        /// <param name="o">The o.</param>
        /// <returns>
        /// 	<c>true</c> if [contains] [the specified o]; otherwise, <c>false</c>.
        /// </returns>
		public bool Contains(object o)
		{
			return impl.ContainsKey(o);
		}

        /// <summary>
        /// Copies the elements of the <see cref="T:System.Collections.ICollection"/> to an <see cref="T:System.Array"/>, starting at a particular <see cref="T:System.Array"/> index.
        /// </summary>
        /// <param name="array">The one-dimensional <see cref="T:System.Array"/> that is the destination of the elements copied from <see cref="T:System.Collections.ICollection"/>. The <see cref="T:System.Array"/> must have zero-based indexing.</param>
        /// <param name="index">The zero-based index in <paramref name="array"/> at which copying begins.</param>
        /// <exception cref="T:System.ArgumentNullException">
        /// 	<paramref name="array"/> is null.
        /// </exception>
        /// <exception cref="T:System.ArgumentOutOfRangeException">
        /// 	<paramref name="index"/> is less than zero.
        /// </exception>
        /// <exception cref="T:System.ArgumentException">
        /// 	<paramref name="array"/> is multidimensional.
        /// -or-
        /// <paramref name="index"/> is equal to or greater than the length of <paramref name="array"/>.
        /// -or-
        /// The number of elements in the source <see cref="T:System.Collections.ICollection"/> is greater than the available space from <paramref name="index"/> to the end of the destination <paramref name="array"/>.
        /// </exception>
        /// <exception cref="T:System.ArgumentException">
        /// The type of the source <see cref="T:System.Collections.ICollection"/> cannot be cast automatically to the type of the destination <paramref name="array"/>.
        /// </exception>
		public void CopyTo(Array array, int index)
		{
			impl.Keys.CopyTo(array, index);
		}

        /// <summary>
        /// Gets the number of elements contained in the <see cref="T:System.Collections.ICollection"/>.
        /// </summary>
        /// <value></value>
        /// <returns>
        /// The number of elements contained in the <see cref="T:System.Collections.ICollection"/>.
        /// </returns>
		public int Count
		{
			get { return impl.Count; }
		}

        /// <summary>
        /// Returns an enumerator that iterates through a collection.
        /// </summary>
        /// <returns>
        /// An <see cref="T:System.Collections.IEnumerator"/> object that can be used to iterate through the collection.
        /// </returns>
		public IEnumerator GetEnumerator()
		{
			return impl.Keys.GetEnumerator();
		}

        /// <summary>
        /// Gets a value indicating whether access to the <see cref="T:System.Collections.ICollection"/> is synchronized (thread safe).
        /// </summary>
        /// <value></value>
        /// <returns>true if access to the <see cref="T:System.Collections.ICollection"/> is synchronized (thread safe); otherwise, false.
        /// </returns>
		public bool IsSynchronized
		{
			get { return impl.IsSynchronized; }
		}

        /// <summary>
        /// Removes the specified o.
        /// </summary>
        /// <param name="o">The o.</param>
		public void Remove(object o)
		{
			impl.Remove(o);
		}

        /// <summary>
        /// Gets an object that can be used to synchronize access to the <see cref="T:System.Collections.ICollection"/>.
        /// </summary>
        /// <value></value>
        /// <returns>
        /// An object that can be used to synchronize access to the <see cref="T:System.Collections.ICollection"/>.
        /// </returns>
		public object SyncRoot
		{
			get { return impl.SyncRoot; }
		}


        /// <summary>
        /// Removes all of the elements from this set.
        /// The set will be empty after this call returns.
        /// </summary>
        public void Clear()
        {
            impl.Clear();
        }

    }
}
