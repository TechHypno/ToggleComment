using System;
using System.Collections.Generic;

namespace ToggleComment.Utils
{
    /// <summary>
    /// <see cref="IDictionary{TKey, TValue}"/>の拡張メソッドを定義するクラスです。
    /// </summary>
    public static class DictionaryExtensions
    {
        /// <summary>
        /// キーがまだ存在しない場合に、<see cref="IDictionary{TKey, TValue}"/>にキーと値のペアを追加します。
        /// </summary>
        /// <typeparam name="TKey">キーの型</typeparam>
        /// <typeparam name="TValue">値の型</typeparam>
        /// <param name="dictionary">呼出し対象の<see cref="IDictionary{TKey, TValue}"/>インスタンス</param>
        /// <param name="key">追加する要素のキー</param>
        /// <param name="addValueFactory">キーの値を生成するために使用される関数</param>
        /// <returns>キーの値。キーが<paramref name="dictionary"/>に既に存在する場合はキーの既存の値、キーが存在していなかった場合は<paramref name="addValueFactory"/>から返されたキーの新しい値になります。</returns>
        /// <exception cref="ArgumentNullException"><paramref name="dictionary"/>または<paramref name="addValueFactory"/>が<see langword="null"/>です。</exception>
        public static TValue GetOrAdd<TKey, TValue>(this IDictionary<TKey, TValue> dictionary, TKey key, bool is2X, Func<TKey, bool, TValue> addValueFactory)
        {
            if (dictionary == null)
            {
                throw new ArgumentNullException(nameof(dictionary));
            }
            if (addValueFactory == null)
            {
                throw new ArgumentNullException(nameof(addValueFactory));
            }

            TValue value;
            if (dictionary.TryGetValue(key, out value) == false)
            {
                value = addValueFactory(key, is2X);
                dictionary.Add(key, value);
            }

            return value;
        }
    }
}
