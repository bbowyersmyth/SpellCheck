using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace PlatformSpellCheck.Tests
{
    [TestClass]
    public class SpellCheckerTests
    {
        [TestMethod]
        public void IsPlatformSupportedTest()
        {
            Assert.IsTrue(SpellChecker.IsPlatformSupported());
        }

        [TestMethod]
        public void IsLanguageSupportedTest()
        {
            Assert.IsTrue(SpellChecker.IsLanguageSupported("en-us"));
            Assert.IsFalse(SpellChecker.IsLanguageSupported("zz-zz"));
        }

        [TestMethod]
        public void SupportedLanguagesTest()
        {
            Assert.IsNotEmpty(SpellChecker.SupportedLanguages);
        }

        [TestMethod]
        public void SuggestionsTest()
        {
            var spell = new SpellChecker();

            var examples = spell.Suggestions("manle");

            Assert.IsNotEmpty(examples);
        }

        [TestMethod]
        public void CheckTest()
        {
            var spell = new SpellChecker("en-us");

            var examples = spell.Check("foxx or teh").ToList();

            Assert.HasCount(2, examples);

            var firstError = examples[0];

            Assert.AreEqual(0, firstError.StartIndex);
            Assert.AreEqual(4, firstError.Length);
            Assert.AreEqual(RecommendedAction.GetSuggestions, firstError.RecommendedAction);
            Assert.AreEqual(string.Empty, firstError.RecommendedReplacement);

            var secondError = examples[1];

            Assert.AreEqual(8, secondError.StartIndex);
            Assert.AreEqual(3, secondError.Length);
            Assert.AreEqual(RecommendedAction.Replace, secondError.RecommendedAction);
            Assert.AreEqual("the", secondError.RecommendedReplacement);
        }

        [TestMethod]
        public void MultiSessionTest()
        {
            var spell1 = new SpellChecker("en-us");
            var spell2 = new SpellChecker("fr-fr");

            var examples1 = spell1.Suggestions("doog").ToList();
            spell1.Dispose();
            var examples2 = spell2.Suggestions("doog").ToList();

            Assert.HasCount(4, examples1);
            Assert.HasCount(2, examples2);
        }

        [TestMethod]
        public void UnsupportedLangTest()
        {
            Assert.ThrowsExactly<ArgumentException>(() => new SpellChecker("zz-zz"));
        }

        [TestMethod]
        public void IgnoreTest()
        {
            var spell = new SpellChecker();

            var examples = spell.Check("codez");

            Assert.IsNotEmpty(examples);

            spell.Ignore("codez");

            examples = spell.Check("codez");

            Assert.IsEmpty(examples);
        }

        [TestMethod]
        public void LanguageIdTest()
        {
            var spell = new SpellChecker("en-AU");

            Assert.AreEqual("en-AU", spell.LanguageTag);
        }

        [TestMethod]
        public void DoubleDisposeTest()
        {
            var spell = new SpellChecker("en-AU");

            spell.Dispose();
            spell.Dispose(); // No error
        }
    }
}
