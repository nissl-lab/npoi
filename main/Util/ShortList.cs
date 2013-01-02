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

namespace NPOI.Util
{
    using System;

    /// <summary>
    /// A List of short's; as full an implementation of the java.Util.List
    /// interface as possible, with an eye toward minimal creation of
    /// objects
    /// 
    /// the mimicry of List is as follows:
    /// <ul>
    /// <li> if possible, operations designated 'optional' in the List
    ///      interface are attempted</li>
    /// <li> wherever the List interface refers to an Object, substitute
    ///      short</li>
    /// <li> wherever the List interface refers to a Collection or List,
    ///      substitute shortList</li>
    /// </ul>
    /// 
    /// the mimicry is not perfect, however:
    /// <ul>
    /// <li> operations involving Iterators or ListIterators are not
    ///      supported</li>
    /// <li> Remove(Object) becomes RemoveValue to distinguish it from
    ///      Remove(short index)</li>
    /// <li> subList is not supported</li>
    /// </ul> 
    /// </summary>
    public class ShortList
    {
        private short[] _array;
        private int _limit;
        private static int _default_size = 128;

        /// <summary>
        /// create an shortList of default size
        /// </summary>
        public ShortList()
            : this(_default_size)
        {

        }

        /// <summary>
        /// create a copy of an existing shortList
        /// </summary>
        /// <param name="list">the existing shortList</param>
        public ShortList(ShortList list)
            : this(list._array.Length)
        {
            Array.Copy(list._array, 0, _array, 0, _array.Length);
            _limit = list._limit;
        }

        /// <summary>
        /// create an shortList with a predefined Initial size
        /// </summary>
        /// <param name="InitialCapacity">the size for the internal array</param>
        public ShortList(int InitialCapacity)
        {
            _array = new short[InitialCapacity];
            _limit = 0;
        }

        /// <summary>
        /// add the specfied value at the specified index
        /// </summary>
        /// <param name="index">the index where the new value is to be Added</param>
        /// <param name="value">the new value</param>
        public void Add(int index, short value)
        {
            if (index > _limit)
            {
                throw new IndexOutOfRangeException();
            }
            else if (index == _limit)
            {
                Add(value);
            }
            else
            {

                // index < limit -- insert into the middle
                if (_limit == _array.Length)
                {
                    GrowArray(_limit * 2);
                }
                Array.Copy(_array, index, _array, index + 1,
                                 _limit - index);
                _array[index] = value;
                _limit++;
            }
        }

        /// <summary>
        /// Appends the specified element to the end of this list
        /// </summary>
        /// <param name="value">element to be Appended to this list.</param>
        /// <returns>return true (as per the general contract of the Collection.add method).</returns>
        public bool Add(short value)
        {
            if (_limit == _array.Length)
            {
                GrowArray(_limit * 2);
            }
            _array[_limit++] = value;
            return true;
        }

        /// <summary>
        /// Appends all of the elements in the specified collection to the
        /// end of this list, in the order that they are returned by the
        /// specified collection's iterator.  The behavior of this
        /// operation is unspecified if the specified collection is
        /// modified while the operation is in progress.  (Note that this
        /// will occur if the specified collection is this list, and it's
        /// nonempty.)
        /// </summary>
        /// <param name="c">collection whose elements are to be Added to this list.</param>
        /// <returns>return true if this list Changed as a result of the call.</returns>
        public bool AddAll(ShortList c)
        {
            if (c._limit != 0)
            {
                if ((_limit + c._limit) > _array.Length)
                {
                    GrowArray(_limit + c._limit);
                }
                Array.Copy(c._array, 0, _array, _limit, c._limit);
                _limit += c._limit;
            }
            return true;
        }

        /// <summary>
        /// Inserts all of the elements in the specified collection into
        /// this list at the specified position.  Shifts the element
        /// currently at that position (if any) and any subsequent elements
        /// to the right (increases their indices).  The new elements will
        /// appear in this list in the order that they are returned by the
        /// specified collection's iterator.  The behavior of this
        /// operation is unspecified if the specified collection is
        /// modified while the operation is in progress.  (Note that this
        /// will occur if the specified collection is this list, and it's
        /// nonempty.)
        /// </summary>
        /// <param name="index">index at which to insert first element from the specified collection.</param>
        /// <param name="c">elements to be inserted into this list.</param>
        /// <returns>return true if this list Changed as a result of the call.</returns>
        /// <exception cref="IndexOutOfRangeException"> if the index is out of range (index &lt; 0 || index &gt; size())</exception>
        public bool AddAll(int index, ShortList c)
        {
            if (index > _limit)
            {
                throw new IndexOutOfRangeException();
            }
            if (c._limit != 0)
            {
                if ((_limit + c._limit) > _array.Length)
                {
                    GrowArray(_limit + c._limit);
                }

                // make a hole
                Array.Copy(_array, index, _array, index + c._limit,
                                 _limit - index);

                // fill it in
                Array.Copy(c._array, 0, _array, index, c._limit);
                _limit += c._limit;
            }
            return true;
        }

        /// <summary>
        /// Removes all of the elements from this list.  This list will be
        /// empty After this call returns (unless it throws an exception).
        /// </summary>
        public void Clear()
        {
            _limit = 0;
        }

        /// <summary>
        /// Returns true if this list Contains the specified element.  More
        /// formally, returns true if and only if this list Contains at
        /// least one element e such that o == e
        /// </summary>
        /// <param name="o">element whose presence in this list is to be Tested.</param>
        /// <returns>return true if this list Contains the specified element.</returns>
        public bool Contains(short o)
        {
            bool rval = false;

            for (int j = 0; !rval && (j < _limit); j++)
            {
                if (_array[j] == o)
                {
                    rval = true;
                }
            }
            return rval;
        }

        /// <summary>
        /// Returns true if this list Contains all of the elements of the specified collection.
        /// </summary>
        /// <param name="c">collection to be Checked for Containment in this list.</param>
        /// <returns>return true if this list Contains all of the elements of the specified collection.</returns>
        public bool ContainsAll(ShortList c)
        {
            bool rval = true;

            if (this != c)
            {
                for (int j = 0; rval && (j < c._limit); j++)
                {
                    if (!Contains(c._array[j]))
                    {
                        rval = false;
                    }
                }
            }
            return rval;
        }

        /// <summary>
        /// Compares the specified object with this list for Equality.
        /// Returns true if and only if the specified object is also a
        /// list, both lists have the same size, and all corresponding
        /// pairs of elements in the two lists are Equal.  (Two elements e1
        /// and e2 are equal if e1 == e2.)  In other words, two lists are
        /// defined to be equal if they contain the same elements in the
        /// same order.  This defInition ensures that the Equals method
        /// works properly across different implementations of the List
        /// interface.
        /// </summary>
        /// <param name="o">the object to be Compared for Equality with this list.</param>
        /// <returns>return true if the specified object is equal to this list.</returns>
        public override bool Equals(Object o)
        {
            bool rval = this == o;

            if (!rval && (o != null) && (o.GetType() == this.GetType()))
            {
                ShortList other = (ShortList)o;

                if (other._limit == _limit)
                {

                    // assume match
                    rval = true;
                    for (int j = 0; rval && (j < _limit); j++)
                    {
                        rval = _array[j] == other._array[j];
                    }
                }
            }
            return rval;
        }

        /// <summary>
        /// Returns the element at the specified position in this list.
        /// </summary>
        /// <param name="index">index of element to return.</param>
        /// <returns>return the element at the specified position in this list.</returns>
        public short Get(int index)
        {
            if (index >= _limit)
            {
                throw new IndexOutOfRangeException();
            }
            return _array[index];
        }

        /// <summary>
        /// Returns the hash code value for this list.  The hash code of a
        /// list is defined to be the result of the following calculation:
        /// 
        /// <code>
        /// hashCode = 1;
        /// Iterator i = list.Iterator();
        /// while (i.HasNext()) {
        ///      Object obj = i.Next();
        ///      hashCode = 31*hashCode + (obj==null ? 0 : obj.HashCode());
        /// }
        /// </code>
        /// 
        /// This ensures that list1.Equals(list2) implies that
        /// list1.HashCode()==list2.HashCode() for any two lists, list1 and
        /// list2, as required by the general contract of Object.HashCode.
        /// </summary>
        /// <returns>return the hash code value for this list.</returns>
        public override int GetHashCode()
        {
            int hash = 0;

            for (int j = 0; j < _limit; j++)
            {
                hash = (31 * hash) + _array[j];
            }
            return hash;
        }

        /// <summary>
        /// Returns the index in this list of the first occurrence of the
        /// specified element, or -1 if this list does not contain this
        /// element.  More formally, returns the lowest index i such that
        /// (o == Get(i)), or -1 if there is no such index.
        /// </summary>
        /// <param name="o">element to search for.</param>
        /// <returns>the index in this list of the first occurrence of the
        /// specified element, or -1 if this list does not contain
        /// this element.
        /// </returns>
        public int IndexOf(short o)
        {
            int rval = 0;

            for (; rval < _limit; rval++)
            {
                if (o == _array[rval])
                {
                    break;
                }
            }
            if (rval == _limit)
            {
                rval = -1;   // didn't find it
            }
            return rval;
        }

        /// <summary>
        /// Returns true if this list Contains no elements.
        /// </summary>
        /// <returns>return true if this list Contains no elements.</returns>
        public bool IsEmpty()
        {
            return _limit == 0;
        }

        /// <summary>
        /// Returns the index in this list of the last occurrence of the
        /// specified element, or -1 if this list does not contain this
        /// element.  More formally, returns the highest index i such that
        /// (o == Get(i)), or -1 if there is no such index.
        /// </summary>
        /// <param name="o">element to search for.</param>
        /// <returns>return the index in this list of the last occurrence of the
        /// specified element, or -1 if this list does not contain this element.</returns>
        public int LastIndexOf(short o)
        {
            int rval = _limit - 1;

            for (; rval >= 0; rval--)
            {
                if (o == _array[rval])
                {
                    break;
                }
            }
            return rval;
        }

        /// <summary>
        /// Removes the element at the specified position in this list.
        /// Shifts any subsequent elements to the left (subtracts one from
        /// their indices).  Returns the element that was Removed from the
        /// list.
        /// </summary>
        /// <param name="index">the index of the element to Removed.</param>
        /// <returns>return the element previously at the specified position.</returns>
        public short Remove(int index)
        {
            if (index >= _limit)
            {
                throw new IndexOutOfRangeException();
            }
            short rval = _array[index];

            Array.Copy(_array, index + 1, _array, index, _limit - index);
            _limit--;
            return rval;
        }

        /// <summary>
        /// Removes the first occurrence in this list of the specified
        /// element (optional operation).  If this list does not contain
        /// the element, it is unChanged.  More formally, Removes the
        /// element with the lowest index i such that (o.Equals(get(i)))
        /// (if such an element exists).
        /// </summary>
        /// <param name="o">element to be Removed from this list, if present.</param>
        /// <returns>return true if this list Contained the specified element.</returns>
        public bool RemoveValue(short o)
        {
            bool rval = false;

            for (int j = 0; !rval && (j < _limit); j++)
            {
                if (o == _array[j])
                {
                    Array.Copy(_array, j + 1, _array, j, _limit - j);
                    _limit--;
                    rval = true;
                }
            }
            return rval;
        }

        /// <summary>
        /// Removes from this list all the elements that are Contained in the specified collection
        /// </summary>
        /// <param name="c">collection that defines which elements will be removed from this list.</param>
        /// <returns>return true if this list Changed as a result of the call.</returns>
        public bool RemoveAll(ShortList c)
        {
            bool rval = false;

            for (int j = 0; j < c._limit; j++)
            {
                if (RemoveValue(c._array[j]))
                {
                    rval = true;
                }
            }
            return rval;
        }

        /// <summary>
        /// Retains only the elements in this list that are Contained in
        /// the specified collection.  In other words, Removes from this
        /// list all the elements that are not Contained in the specified
        /// collection.
        /// </summary>
        /// <param name="c">collection that defines which elements this Set will retain.</param>
        /// <returns>return true if this list Changed as a result of the call.</returns>
        public bool RetainAll(ShortList c)
        {
            bool rval = false;

            for (int j = 0; j < _limit; )
            {
                if (!c.Contains(_array[j]))
                {
                    Remove(j);
                    rval = true;
                }
                else
                {
                    j++;
                }
            }
            return rval;
        }

        /// <summary>
        /// Replaces the element at the specified position in this list with the specified element
        /// </summary>
        /// <param name="index">index of element to Replace.</param>
        /// <param name="element">element to be stored at the specified position.</param>
        /// <returns>return the element previously at the specified position.</returns>
        public short Set(int index, short element)
        {
            if (index >= _limit)
            {
                throw new IndexOutOfRangeException();
            }
            short rval = _array[index];

            _array[index] = element;
            return rval;
        }

        /// <summary>
        /// Returns the number of elements in this list. If this list
        /// Contains more than Int32.MaxValue elements, returns
        /// Int32.MaxValue.
        /// </summary>
        /// <returns>return the number of elements in this shortList</returns>
        public int Size()
        {
            return _limit;
        }
        /// <summary>
        /// the number of elements in this shortList
        /// </summary>
        public int Count
        {
            get { return _limit; }
        }
        
        /// <summary>
        /// Returns an array Containing all of the elements in this list in
        /// proper sequence.  Obeys the general contract of the
        /// Collection.ToArray method.
        /// </summary>
        /// <returns>an array Containing all of the elements in this list in
        /// proper sequence.</returns>
        public short[] ToArray()
        {
            short[] rval = new short[_limit];

            Array.Copy(_array, 0, rval, 0, _limit);
            return rval;
        }

        /// <summary>
        /// Returns an array Containing all of the elements in this list in
        /// proper sequence.  Obeys the general contract of the
        /// Collection.ToArray(Object[]) method.
        /// </summary>
        /// <param name="a">the array into which the elements of this list are to
        /// be stored, if it is big enough; otherwise, a new array
        /// is allocated for this purpose.</param>
        /// <returns>return an array Containing the elements of this list.</returns>
        public short[] ToArray(short[] a)
        {
            short[] rval;

            if (a.Length == _limit)
            {
                Array.Copy(_array, 0, a, 0, _limit);
                rval = a;
            }
            else
            {
                rval = ToArray();
            }
            return rval;
        }

        private void GrowArray(int new_size)
        {
            int size = (new_size == _array.Length) ? new_size + 1
                                                            : new_size;
            short[] new_array = new short[size];

            Array.Copy(_array, 0, new_array, 0, _limit);
            _array = new_array;
        }
    }   // end public class shortList
}