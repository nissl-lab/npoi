namespace System.Collections.Generic
{
    public interface IBidirectionalDictionary<TKey, TValue> : IDictionary<TKey, TValue>
        where TKey : notnull
        where TValue : notnull
    {
        public IBidirectionalDictionary<TValue, TKey> Inverse { get; }
        public bool ContainsValue(TValue value);
    }
}
