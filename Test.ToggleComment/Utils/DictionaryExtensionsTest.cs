using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ToggleComment.Codes;
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
                ["key1"] = "Hoge",
                ["key2"] = "Rambo"
            };

            Assert.AreEqual("Hoge", stringValues.GetOrAdd("key1", false, (key, is2X) => key + "_Fuga"));
            Assert.AreEqual("Hoge", stringValues["key1"]);

            Assert.AreEqual("Rambo", stringValues.GetOrAdd("key2", false, (key, is2X) => key + "_Piyo"));
            Assert.AreEqual("Rambo", stringValues["key2"]);
        }
    }
}
