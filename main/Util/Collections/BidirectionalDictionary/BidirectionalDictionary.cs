#nullable enable
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;

namespace System.Collections.Generic
{
    /// <summary>
    /// Represents a dictionary with non-null unique values that provides access to an inverse dictionary.
    /// </summary>
    /// <typeparam name="TKey">The type of the keys in the dictionary.</typeparam>
    /// <typeparam name="TValue">The type of the values in the dictionary.</typeparam>
    [DebuggerTypeProxy(typeof(DictionaryDebugView<,>))]
    [DebuggerDisplay("Count = {Count}")]
    public class BidirectionalDictionary<TKey, TValue> : IBidirectionalDictionary<TKey, TValue>, IReadOnlyBidirectionalDictionary<TKey, TValue>
        where TKey : notnull
        where TValue : notnull
    {
        protected internal readonly Dictionary<TKey, TValue> _baseDictionary;

        #region Properties

        /// <summary>
        /// Gets the inverse <see cref="BidirectionalDictionary{TKey,TValue}"/>.
        /// </summary>
        public BidirectionalDictionary<TValue, TKey> Inverse { get; }

        /// <summary>
        /// Gets the <see cref="IEqualityComparer{T}"/> that is used to determine equality of keys for the dictionary.
        /// </summary>
        public IEqualityComparer<TKey> KeyComparer => _baseDictionary.Comparer;

        /// <summary>
        /// Gets the <see cref="IEqualityComparer{T}"/> that is used to determine equality of values for the dictionary.
        /// </summary>
        public IEqualityComparer<TValue> ValueComparer { get; }

        /// <summary>
        /// Gets the number of key/value pairs contained in the <see cref="BidirectionalDictionary{TKey, TValue}"/>.
        /// </summary>
        public int Count => _baseDictionary.Count;

        /// <summary>
        /// Gets a collection containing the keys in the <see cref="BidirectionalDictionary{TKey, TValue}"/>.
        /// </summary>
        public Dictionary<TKey, TValue>.KeyCollection Keys => _baseDictionary.Keys;

        /// <summary>
        /// Gets a collection containing the values in the <see cref="BidirectionalDictionary{TKey, TValue}"/>.
        /// </summary>
        public Dictionary<TKey, TValue>.ValueCollection Values => _baseDictionary.Values;

        /// <summary>
        /// Gets or sets the value associated with the specified key.
        /// </summary>
        /// <param name="key">The key of the value to get or set.</param>
        /// <returns>The value associated with the specified key. If the specified key is not found, a get operation throws a
        /// <see cref="KeyNotFoundException"/>, and a set operation creates a new element with the specified key.</returns>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException"></exception>
        /// <exception cref="KeyNotFoundException"></exception>
        public TValue this[TKey key]
        {
            get => _baseDictionary[key];
            set
            {
                if (key == null)
                    throw new ArgumentNullException(nameof(key));

                if (value == null)
                    throw new ArgumentNullException(nameof(value));

                if (TryGetValue(key, out var oldValue))
                {
                    if (ValueComparer.Equals(oldValue, value))
                        return;

                    if (ContainsValue(value))
                    {
                        throw new ArgumentException("The value already exists.", nameof(value));
                    }
                    else
                    {
                        _baseDictionary[key] = value;

                        Inverse._baseDictionary.Remove(oldValue);
                        Inverse._baseDictionary.Add(value, key);
                    }
                }
                else
                {
                    Add(key, value);
                }
            }
        }

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        IBidirectionalDictionary<TValue, TKey> IBidirectionalDictionary<TKey, TValue>.Inverse => Inverse;

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        IReadOnlyBidirectionalDictionary<TValue, TKey> IReadOnlyBidirectionalDictionary<TKey, TValue>.Inverse => Inverse.AsReadOnly();

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        ICollection<TKey> IDictionary<TKey, TValue>.Keys => Keys;
        
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        ICollection<TValue> IDictionary<TKey, TValue>.Values => Values;
        
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        IEnumerable<TKey> IReadOnlyDictionary<TKey, TValue>.Keys => Keys;
        
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        IEnumerable<TValue> IReadOnlyDictionary<TKey, TValue>.Values => Values;
        
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        bool ICollection<KeyValuePair<TKey, TValue>>.IsReadOnly => false;

        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="BidirectionalDictionary{TKey, TValue}"/> class that is empty,
        /// has the default initial capacity, and uses the default equality comparers.
        /// </summary>
        public BidirectionalDictionary() : this(new Dictionary<TKey, TValue>()) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BidirectionalDictionary{TKey, TValue}"/> class that is empty,
        /// has the specified initial capacity, and uses the default equality comparers.
        /// </summary>
        /// <param name="capacity">The initial number of elements that <see cref="BidirectionalDictionary{TKey, TValue}"/> can contain.</param>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        public BidirectionalDictionary(int capacity) : this(capacity, null, null) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BidirectionalDictionary{TKey, TValue}"/> class
        /// that contains elements copied from the specified <see cref="IDictionary{TKey, TValue}"/>
        /// and uses the default equality comparers.
        /// </summary>
        /// <param name="dictionary">The <see cref="IDictionary{TKey, TValue}"/> whose elements are copied to the new <see cref="BidirectionalDictionary{TKey, TValue}"/>.</param>
        /// <exception cref="ArgumentNullException"></exception>
        public BidirectionalDictionary(IDictionary<TKey, TValue> dictionary)
            : this(new Dictionary<TKey, TValue>(dictionary)) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BidirectionalDictionary{TKey, TValue}"/> class
        /// that contains elements copied from the specified <see cref="IEnumerable{T}"/>
        /// and uses the default equality comparers.
        /// </summary>
        /// <param name="collection">The <see cref="IEnumerable{T}"/> whose elements are copied to the new <see cref="BidirectionalDictionary{TKey, TValue}"/>.</param>
        /// <exception cref="ArgumentNullException"></exception>

#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
        public BidirectionalDictionary(IEnumerable<KeyValuePair<TKey, TValue>> collection)
            : this(new Dictionary<TKey, TValue>(collection)) { }
#elif NETSTANDARD2_0 || NET472 || NET40 || NET45
        public BidirectionalDictionary(IEnumerable<KeyValuePair<TKey, TValue>> collection)
            : this(new Dictionary<TKey, TValue>(collection?.ToDictionary(pair => pair.Key, pair => pair.Value))) { }
#endif

        /// <summary>
        /// Initializes a new instance of the <see cref="BidirectionalDictionary{TKey, TValue}"/> class that is empty,
        /// has the default initial capacity, and uses the specified equality comparers.
        /// </summary>
        /// <param name="keyComparer">The <see cref="IEqualityComparer{T}"/> implementation to use when
        /// comparing keys, or null to use the default <see cref="IEqualityComparer{T}"/> for the type of the key.</param>
        /// <param name="valueComparer">The <see cref="IEqualityComparer{T}"/> implementation to use when
        /// comparing values, or null to use the default <see cref="IEqualityComparer{T}"/> for the type of the value.</param>
        public BidirectionalDictionary(IEqualityComparer<TKey>? keyComparer, IEqualityComparer<TValue>? valueComparer)
            : this(new Dictionary<TKey, TValue>(keyComparer), valueComparer) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BidirectionalDictionary{TKey, TValue}"/> class that is empty,
        /// has the specified initial capacity, and uses the specified equality comparers.
        /// </summary>
        /// <param name="capacity">The initial number of elements that <see cref="BidirectionalDictionary{TKey, TValue}"/> can contain.</param>
        /// <param name="keyComparer">The <see cref="IEqualityComparer{T}"/> implementation to use when
        /// comparing keys, or null to use the default <see cref="IEqualityComparer{T}"/> for the type of the key.</param>
        /// <param name="valueComparer">The <see cref="IEqualityComparer{T}"/> implementation to use when
        /// comparing values, or null to use the default <see cref="IEqualityComparer{T}"/> for the type of the value.</param>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        public BidirectionalDictionary(int capacity, IEqualityComparer<TKey>? keyComparer, IEqualityComparer<TValue>? valueComparer)
            : this(new Dictionary<TKey, TValue>(capacity, keyComparer), valueComparer) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BidirectionalDictionary{TKey, TValue}"/> class that
        /// contains elements copied from the specified <see cref="IDictionary{TKey, TValue}"/>, and uses the specified equality comparers.
        /// </summary>
        /// <param name="dictionary">The <see cref="IDictionary{TKey, TValue}"/> whose elements are copied to the new <see cref="BidirectionalDictionary{TKey, TValue}"/>.</param>
        /// <param name="keyComparer">The <see cref="IEqualityComparer{T}"/> implementation to use when
        /// comparing keys, or null to use the default <see cref="IEqualityComparer{T}"/> for the type of the key.</param>
        /// <param name="valueComparer">The <see cref="IEqualityComparer{T}"/> implementation to use when
        /// comparing values, or null to use the default <see cref="IEqualityComparer{T}"/> for the type of the value.</param>
        /// <exception cref="ArgumentNullException"></exception>
        public BidirectionalDictionary(IDictionary<TKey, TValue> dictionary,
            IEqualityComparer<TKey>? keyComparer, IEqualityComparer<TValue>? valueComparer)
            : this(new Dictionary<TKey, TValue>(dictionary, keyComparer), valueComparer) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BidirectionalDictionary{TKey, TValue}"/> class that
        /// contains elements copied from the specified <see cref="IEnumerable{T}"/>, and uses the specified equality comparers.
        /// </summary>
        /// <param name="collection">The <see cref="IEnumerable{T}"/> whose elements are copied to the new <see cref="BidirectionalDictionary{TKey, TValue}"/>.</param>
        /// <param name="keyComparer">The <see cref="IEqualityComparer{T}"/> implementation to use when
        /// comparing keys, or null to use the default <see cref="IEqualityComparer{T}"/> for the type of the key.</param>
        /// <param name="valueComparer">The <see cref="IEqualityComparer{T}"/> implementation to use when
        /// comparing values, or null to use the default <see cref="IEqualityComparer{T}"/> for the type of the value.</param>
        /// <exception cref="ArgumentNullException"></exception>

#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
        public BidirectionalDictionary(IEnumerable<KeyValuePair<TKey, TValue>> collection,
            IEqualityComparer<TKey>? keyComparer, IEqualityComparer<TValue>? valueComparer)
            : this(new Dictionary<TKey, TValue>(collection, keyComparer), valueComparer) { }
#elif NETSTANDARD2_0 || NET40 || NET45 || NET472
        public BidirectionalDictionary(IEnumerable<KeyValuePair<TKey, TValue>> collection,
            IEqualityComparer<TKey>? keyComparer, IEqualityComparer<TValue>? valueComparer)
            : this(new Dictionary<TKey, TValue>(collection?.ToDictionary(pair => pair.Key, pair => pair.Value, keyComparer)),
                  valueComparer) { }
#endif

        private BidirectionalDictionary(BidirectionalDictionary<TValue, TKey> inverse)
        {
            _baseDictionary = inverse._baseDictionary.ToDictionary(pair => pair.Value, pair => pair.Key, inverse.ValueComparer);
            ValueComparer   = inverse.KeyComparer;
            Inverse         = inverse;
        }

        private BidirectionalDictionary(Dictionary<TKey, TValue> dictionary, IEqualityComparer<TValue>? valueComparer = null)
        {
            _baseDictionary = dictionary;
            ValueComparer   = valueComparer ?? EqualityComparer<TValue>.Default;
            Inverse         = new BidirectionalDictionary<TValue, TKey>(this);
        }

        #endregion

        #region Methods

        /// <summary>
        /// Adds the specified key and value to the dictionary.
        /// </summary>
        /// <param name="key">The key of the element to add.</param>
        /// <param name="value">The value of the element to add.</param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException"></exception>
        public void Add(TKey key, TValue value)
        {
            if (key == null)
                throw new ArgumentNullException(nameof(key));

            if (value == null)
                throw new ArgumentNullException(nameof(value));

            if (ContainsKey(key))
                throw new ArgumentException("The same key already exists.", nameof(key));

            if (ContainsValue(value))
                throw new ArgumentException("The same value already exists.", nameof(value));

            _baseDictionary.Add(key, value);
            Inverse._baseDictionary.Add(value, key);
        }

        /// <summary>
        /// Removes the value with the specified key from the <see cref="BidirectionalDictionary{TKey, TValue}"/>.
        /// </summary>
        /// <param name="key">The key of the element to remove.</param>
        /// <returns><see langword="true"/> if the element is successfully found and removed; otherwise, <see langword="false"/>.
        /// This method returns <see langword="false"/> if key is not found in the <see cref="BidirectionalDictionary{TKey, TValue}"/>.</returns>
        /// <exception cref="ArgumentNullException"></exception>
        public bool Remove(TKey key) => Remove(key, out _);

        /// <summary>
        /// Removes the value with the specified key from the <see cref="BidirectionalDictionary{TKey, TValue}"/>.
        /// </summary>
        /// <param name="key">The key of the element to remove.</param>
        /// <param name="value">When this method returns, contains the value associated with the specified key,
        /// if the key is found; otherwise, the default value for the type of the value parameter.
        /// This parameter is passed uninitialized.</param>
        /// <returns><see langword="true"/> if the element is successfully found and removed; otherwise, <see langword="false"/>.
        /// This method returns <see langword="false"/> if key is not found in the <see cref="BidirectionalDictionary{TKey, TValue}"/>.</returns>
        /// <exception cref="ArgumentNullException"></exception>
        public bool Remove(TKey key, out TValue value)
        {
            if (key == null)
                throw new ArgumentNullException(nameof(key));
#nullable disable
            return _baseDictionary.TryGetValue(key, out value) &&
                   _baseDictionary.Remove(key) &&
                   Inverse._baseDictionary.Remove(value);
#nullable enable
        }

        /// <summary>
        /// Removes all keys and values from the <see cref="BidirectionalDictionary{TKey, TValue}"/>.
        /// </summary>
        public void Clear()
        {
            _baseDictionary.Clear();
            Inverse._baseDictionary.Clear();
        }

        /// <summary>
        /// Determines whether the <see cref="BidirectionalDictionary{TKey, TValue}"/> contains the specified key.
        /// </summary>
        /// <param name="key">The key to locate in the <see cref="BidirectionalDictionary{TKey, TValue}"/>.</param>
        /// <returns><see langword="true"/> if the <see cref="BidirectionalDictionary{TKey, TValue}"/> contains
        /// an element with the specified key; otherwise, <see langword="false"/>.</returns>
        public bool ContainsKey(TKey key) => _baseDictionary.ContainsKey(key);

        /// <summary>
        /// Determines whether the <see cref="BidirectionalDictionary{TKey, TValue}"/> contains the specified value.
        /// </summary>
        /// <param name="value">The value to locate in the <see cref="BidirectionalDictionary{TKey, TValue}"/>.</param>
        /// <returns><see langword="true"/> if the <see cref="BidirectionalDictionary{TKey, TValue}"/> contains
        /// an element with the specified value; otherwise, <see langword="false"/>.</returns>
        public bool ContainsValue(TValue value) => Inverse.ContainsKey(value);

        /// <summary>
        /// Attempts to add the specified key and value to the <see cref="BidirectionalDictionary{TKey, TValue}"/>.
        /// </summary>
        /// <param name="key">The key of the element to add.</param>
        /// <param name="value">The value of the element to add.</param>
        /// <returns><see langword="true"/> if the key/value pair was added to the <see cref="BidirectionalDictionary{TKey, TValue}"/>
        /// successfully; otherwise, <see langword="false"/>.</returns>
        /// <exception cref="ArgumentNullException"></exception>
        public bool TryAdd(TKey key, TValue value)
        {
            if (key == null)
                throw new ArgumentNullException(nameof(key));

            if (value == null)
                throw new ArgumentNullException(nameof(value));

            if (ContainsKey(key) || ContainsValue(value))
                return false;

            _baseDictionary.Add(key, value);
            Inverse._baseDictionary.Add(value, key);

            return true;
        }

        /// <summary>
        /// Gets the value associated with the specified key.
        /// </summary>
        /// <param name="key">The key of the value to get.</param>
        /// <param name="value">When this method returns, contains the value associated with the specified key,
        /// if the key is found; otherwise, the default value for the type of the value parameter.
        /// This parameter is passed uninitialized.</param>
        /// <returns><see langword="true"/> if the <see cref="BidirectionalDictionary{TKey, TValue}"/> contains
        /// an element with the specified key; otherwise, <see langword="false"/>.</returns>
#nullable disable
        public bool TryGetValue(TKey key, out TValue value) => _baseDictionary.TryGetValue(key, out value);
#nullable enable

        public TKey GetKey(TValue value)
        {
            return this.Inverse[value];
        }

        public TKey RemoveValue(TValue value)
        {
            TKey key = GetKey(value);
            Remove(key);
            return key;
        }

        public TKey LastKey()
        {
            return this.Keys.Last();
        }

        /*public void EnsureCapacity(int capacity)
        {
            _baseDictionary.EnsureCapacity(capacity);
            Inverse._baseDictionary.EnsureCapacity(capacity);
        }

        public void TrimExcess()
        {
            _baseDictionary.TrimExcess();
            Inverse._baseDictionary.TrimExcess();
        }

        public void TrimExcess(int capacity)
        {
            _baseDictionary.TrimExcess(capacity);
            Inverse._baseDictionary.TrimExcess(capacity);
        }*/

        /// <summary>
        /// Returns a read-only <see cref="ReadOnlyBidirectionalDictionary{TKey, TValue}"></see> wrapper for the current dictionary.
        /// </summary>
        /// <returns>An object that acts as a read-only wrapper around the current <see cref="BidirectionalDictionary{TKey, TValue}"></see>.</returns>
        public ReadOnlyBidirectionalDictionary<TKey, TValue> AsReadOnly() => new ReadOnlyBidirectionalDictionary<TKey, TValue>(this);

        public IEnumerator GetEnumerator() => _baseDictionary.GetEnumerator();

        void ICollection<KeyValuePair<TKey, TValue>>.Add(KeyValuePair<TKey, TValue> item) => Add(item.Key, item.Value);

        bool ICollection<KeyValuePair<TKey, TValue>>.Remove(KeyValuePair<TKey, TValue> item)
        {
            if (item.Key == null)
                throw new ArgumentNullException("The item key == null.", nameof(item));

            if (item.Value == null)
                throw new ArgumentNullException("The item value == null.", nameof(item));

            return ((ICollection<KeyValuePair<TKey, TValue>>)_baseDictionary).Remove(item) &&
                Inverse._baseDictionary.Remove(item.Value);
        }

        bool ICollection<KeyValuePair<TKey, TValue>>.Contains(KeyValuePair<TKey, TValue> item)
        {
            if (item.Key == null)
                throw new ArgumentNullException("The item key == null.", nameof(item));

            if (item.Value == null)
                throw new ArgumentNullException("The item value == null.", nameof(item));

            return ((ICollection<KeyValuePair<TKey, TValue>>)_baseDictionary).Contains(item);
        }

        void ICollection<KeyValuePair<TKey, TValue>>.CopyTo(KeyValuePair<TKey, TValue>[] array, int arrayIndex) =>
            ((ICollection<KeyValuePair<TKey, TValue>>)_baseDictionary).CopyTo(array, arrayIndex);

        IEnumerator<KeyValuePair<TKey, TValue>> IEnumerable<KeyValuePair<TKey, TValue>>.GetEnumerator()
            => _baseDictionary.GetEnumerator();

#endregion
    }
}
