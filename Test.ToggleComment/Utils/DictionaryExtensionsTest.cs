using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ToggleComment.Utils;

namespace Test.ToggleComment.Utils
{
    /// <summary>
    /// <see cref="DictionaryExtensions"/>のテストクラスです。
    /// </summary>
    [TestClass]
    public class DictionaryExtensionsTest
    {
        [TestMethod]
        public void GetOrAddTest()
        {
            var stringValues = new Dictionary<string, string>
            {
                ["key1"] = "Hoge"
            };

            Assert.AreEqual("Hoge", stringValues.GetOrAdd("key1", key => "Fuga"));
            Assert.AreEqual("Hoge", stringValues["key1"]);

            Assert.AreEqual("key2_Piyo", stringValues.GetOrAdd("key2", key => key + "_Piyo"));
            Assert.AreEqual("key2_Piyo", stringValues["key2"]);
        }
    }
}