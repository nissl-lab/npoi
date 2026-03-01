using System;

namespace NPOI.Util.Optional
{
    public struct ValueOption<T> : IEquatable<ValueOption<T>> where T : struct
    {
        private T? _content;

        public static ValueOption<T> Some(T obj) => new() { _content = obj };
        public static ValueOption<T> None() => new();

        public Option<TResult> Map<TResult>(Func<T, TResult> map) where TResult : class =>
            _content.HasValue ? Option<TResult>.Some(map(_content.Value)) : Option<TResult>.None();
        public ValueOption<TResult> MapValue<TResult>(Func<T, TResult> map) where TResult : struct =>
            new() { _content = _content.HasValue ? map(_content.Value) : null };

        public Option<TResult> MapOptional<TResult>(Func<T, Option<TResult>> map) where TResult : class =>
            _content.HasValue ? map(_content.Value) : Option<TResult>.None();
        public ValueOption<TResult> MapOptionalValue<TResult>(Func<T, ValueOption<TResult>> map) where TResult : struct =>
            _content.HasValue ? map(_content.Value) : ValueOption<TResult>.None();
        
        public ValueOption<T> IfPresent(Action<T> action)
        {
            if(_content.HasValue) action(_content.Value);
            return this;
        }

        public T? OrElse(T? orElse) => _content ?? orElse;
        public T OrElse(T orElse) => _content ?? orElse;
        public T Reduce(T orElse) => _content ?? orElse;
        public T Reduce(Func<T> orElse) => _content ?? orElse();

        public ValueOption<T> Where(Func<T, bool> predicate) =>
            _content.HasValue && predicate(_content.Value) ? this : ValueOption<T>.None();

        public ValueOption<T> WhereNot(Func<T, bool> predicate) =>
            _content.HasValue && !predicate(_content.Value) ? this : ValueOption<T>.None();

        public override int GetHashCode() => _content?.GetHashCode() ?? 0;
        #nullable enable
        public override bool Equals(object? other) => other is ValueOption<T> option && Equals(option);
        #nullable disable
        public bool Equals(ValueOption<T> other) =>
            _content.HasValue ? other._content.HasValue && _content.Value.Equals(other._content.Value)
            : !other._content.HasValue;

        public static bool operator ==(ValueOption<T> a, ValueOption<T> b) => a.Equals(b);
        public static bool operator !=(ValueOption<T> a, ValueOption<T> b) => !(a.Equals(b));
    }
}