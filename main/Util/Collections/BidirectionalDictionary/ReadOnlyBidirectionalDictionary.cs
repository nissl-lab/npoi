using System.Collections.Generic;
using System.Diagnostics;

namespace System.Collections.ObjectModel
{
    /// <summary>
    /// Represents a read-only dictionary with non-null unique values that provides access to an inverse read-only dictionary.
    /// </summary>
    /// <typeparam name="TKey">The type of the keys in the dictionary.</typeparam>
    /// <typeparam name="TValue">The type of the values in the dictionary.</typeparam>
    [DebuggerTypeProxy(typeof(DictionaryDebugView<,>))]
    [DebuggerDisplay("Count = {Count}")]
    public class ReadOnlyBidirectionalDictionary<TKey, TValue> : IBidirectionalDictionary<TKey, TValue>, IReadOnlyBidirectionalDictionary<TKey, TValue>
        where TKey : notnull
        where TValue : notnull
    {
        private readonly IBidirectionalDictionary<TKey, TValue> _baseDictionary;

        #region Properties

        /// <summary>
        /// Gets the inverse <see cref="ReadOnlyBidirectionalDictionary{TKey,TValue}"/>.
        /// </summary>
        public ReadOnlyBidirectionalDictionary<TValue, TKey> Inverse { get; }

        /// <summary>
        /// Gets the number of key/value pairs contained in the <see cref="ReadOnlyBidirectionalDictionary{TKey, TValue}"/>.
        /// </summary>
        public int Count => _baseDictionary.Count;

        /// <summary>
        /// Gets a collection containing the keys in the <see cref="ReadOnlyBidirectionalDictionary{TKey, TValue}"/>.
        /// </summary>
        public ICollection<TKey> Keys => _baseDictionary.Keys;

        /// <summary>
        /// Gets a collection containing the values in the <see cref="ReadOnlyBidirectionalDictionary{TKey, TValue}"/>.
        /// </summary>
        public ICollection<TValue> Values => _baseDictionary.Values;

        /// <summary>
        /// Gets the value associated with the specified key.
        /// </summary>
        /// <param name="key">The key of the value to get.</param>
        /// <returns>The value associated with the specified key. If the specified key is not found, a get operation throws a
        /// <see cref="KeyNotFoundException"/>.</returns>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="KeyNotFoundException"></exception>
        public TValue this[TKey key] => _baseDictionary[key];

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        IBidirectionalDictionary<TValue, TKey> IBidirectionalDictionary<TKey, TValue>.Inverse => Inverse;

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        IReadOnlyBidirectionalDictionary<TValue, TKey> IReadOnlyBidirectionalDictionary<TKey, TValue>.Inverse => Inverse;

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        TValue IDictionary<TKey, TValue>.this[TKey key]
        {
            get => this[key];
            set => throw new NotSupportedException();
        }

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        ICollection<TKey> IDictionary<TKey, TValue>.Keys => Keys;
        
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        ICollection<TValue> IDictionary<TKey, TValue>.Values => Values;

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        IEnumerable<TKey> IReadOnlyDictionary<TKey, TValue>.Keys => Keys;
        
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        IEnumerable<TValue> IReadOnlyDictionary<TKey, TValue>.Values => Values;
        
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        bool ICollection<KeyValuePair<TKey, TValue>>.IsReadOnly => true;

        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ReadOnlyBidirectionalDictionary{TKey, TValue}"/> class
        /// that is a wrapper around the specified bidirectional dictionary.
        /// </summary>
        /// <param name="bidirectionalDictionary">The bidirectional dictionary to wrap.</param>
        /// <exception cref="ArgumentNullException"></exception>
        public ReadOnlyBidirectionalDictionary(IBidirectionalDictionary<TKey, TValue> bidirectionalDictionary)
        {
            _baseDictionary = bidirectionalDictionary ?? throw new ArgumentNullException(nameof(bidirectionalDictionary));
            Inverse         = new ReadOnlyBidirectionalDictionary<TValue, TKey>(this);
        }

        private ReadOnlyBidirectionalDictionary(ReadOnlyBidirectionalDictionary<TValue, TKey> inverse)
        {
            _baseDictionary = inverse._baseDictionary.Inverse;
            Inverse         = inverse;
        }

        #endregion

        #region Methods

        /// <summary>
        /// Determines whether the <see cref="ReadOnlyBidirectionalDictionary{TKey, TValue}"/> contains the specified key.
        /// </summary>
        /// <param name="key">The key to locate in the <see cref="ReadOnlyBidirectionalDictionary{TKey, TValue}"/>.</param>
        /// <returns><see langword="true"/> if the <see cref="ReadOnlyBidirectionalDictionary{TKey, TValue}"/> contains
        /// an element with the specified key; otherwise, <see langword="false"/>.</returns>
        public bool ContainsKey(TKey key) => _baseDictionary.ContainsKey(key);

        /// <summary>
        /// Determines whether the <see cref="ReadOnlyBidirectionalDictionary{TKey, TValue}"/> contains the specified value.
        /// </summary>
        /// <param name="value">The value to locate in the <see cref="ReadOnlyBidirectionalDictionary{TKey, TValue}"/>.</param>
        /// <returns><see langword="true"/> if the <see cref="ReadOnlyBidirectionalDictionary{TKey, TValue}"/> contains
        /// an element with the specified value; otherwise, <see langword="false"/>.</returns>
        public bool ContainsValue(TValue value) => _baseDictionary.ContainsValue(value);

        /// <summary>
        /// Gets the value associated with the specified key.
        /// </summary>
        /// <param name="key">The key of the value to get.</param>
        /// <param name="value">When this method returns, contains the value associated with the specified key,
        /// if the key is found; otherwise, the default value for the type of the value parameter.
        /// This parameter is passed uninitialized.</param>
        /// <returns><see langword="true"/> if the <see cref="ReadOnlyBidirectionalDictionary{TKey, TValue}"/> contains
        /// an element with the specified key; otherwise, <see langword="false"/>.</returns>
        public bool TryGetValue(TKey key, out TValue value) => _baseDictionary.TryGetValue(key, out value);

        public IEnumerator GetEnumerator() => _baseDictionary.GetEnumerator();

        void IDictionary<TKey, TValue>.Add(TKey key, TValue value) => throw new NotSupportedException();

        bool IDictionary<TKey, TValue>.Remove(TKey key) => throw new NotSupportedException();

        void ICollection<KeyValuePair<TKey, TValue>>.Add(KeyValuePair<TKey, TValue> item) => throw new NotSupportedException();

        bool ICollection<KeyValuePair<TKey, TValue>>.Remove(KeyValuePair<TKey, TValue> item) => throw new NotSupportedException();

        void ICollection<KeyValuePair<TKey, TValue>>.Clear() => throw new NotSupportedException();

        bool ICollection<KeyValuePair<TKey, TValue>>.Contains(KeyValuePair<TKey, TValue> item) =>
            ((ICollection<KeyValuePair<TKey, TValue>>)_baseDictionary).Contains(item);

        void ICollection<KeyValuePair<TKey, TValue>>.CopyTo(KeyValuePair<TKey, TValue>[] array, int arrayIndex) =>
            ((ICollection<KeyValuePair<TKey, TValue>>)_baseDictionary).CopyTo(array, arrayIndex);

        IEnumerator<KeyValuePair<TKey, TValue>> IEnumerable<KeyValuePair<TKey, TValue>>.GetEnumerator()
            => ((IEnumerable<KeyValuePair<TKey, TValue>>)_baseDictionary).GetEnumerator();

        #endregion
    }
}
