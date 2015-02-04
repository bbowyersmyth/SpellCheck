using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace Text.SpellCheck.Tests
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
            Assert.IsTrue(SpellChecker.SupportedLanguages.Count() > 0);
        }

        [TestMethod]
        public void SuggestionsTest()
        {
            var spell = new SpellChecker();

            var examples = spell.Suggestions("manle");

            Assert.IsTrue(examples.Count() > 0);
        }

        [TestMethod]
        public void CheckTest()
        {
            var spell = new SpellChecker("en-us");

            var examples = spell.Check("foxx or recieve").ToList();

            Assert.AreEqual(examples.Count(), 2);

            var firstError = examples[0];

            Assert.AreEqual(firstError.StartIndex, 0);
            Assert.AreEqual(firstError.Length, 4);
            Assert.AreEqual(firstError.RecommendedAction, RecommendedAction.GetSuggestions);
            Assert.AreEqual(firstError.RecommendedReplacement, string.Empty);

            var secondError = examples[1];

            Assert.AreEqual(secondError.StartIndex, 8);
            Assert.AreEqual(secondError.Length, 7);
            Assert.AreEqual(secondError.RecommendedAction, RecommendedAction.Replace);
            Assert.AreEqual(secondError.RecommendedReplacement, "receive");
        }

        [TestMethod]
        public void MultiSessionTest()
        {
            var spell1 = new SpellChecker("en-us");
            var spell2 = new SpellChecker("fr-fr");

            var examples1 = spell1.Suggestions("doog").ToList();
            var examples2 = spell2.Suggestions("doog").ToList();

            Assert.AreEqual(examples1.Count(), 4);
            Assert.AreEqual(examples2.Count(), 2);
        }

        [TestMethod, ExpectedException(typeof(ArgumentException))]
        public void UnsupportedLangTest()
        {
            var spell = new SpellChecker("zz-zz");
        }
    }
}
