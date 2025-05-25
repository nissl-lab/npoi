using System;

namespace NPOI.Util.Optional
{
    public static class OptionalExtensions
    {
        #nullable enable
        public static Option<T> ToOption<T>(this T? obj) where T : class =>
            obj is null ? Option<T>.None() : Option<T>.Some(obj);

        public static Option<T> Where<T>(this T? obj, Func<T, bool> predicate) where T : class =>
            obj is not null && predicate(obj) ? Option<T>.Some(obj) : Option<T>.None();

        public static Option<T> WhereNot<T>(this T? obj, Func<T, bool> predicate) where T : class =>
            obj is not null && !predicate(obj) ? Option<T>.Some(obj) : Option<T>.None();
        #nullable disable
    }
}
