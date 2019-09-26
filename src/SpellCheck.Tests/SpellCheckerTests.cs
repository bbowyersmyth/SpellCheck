// <copyright file="SpellCheckerTests.cs" company="Bruce Bowyer-Smyth">
//     Copyright © Bruce Bowyer-Smyth. All rights reserved.
// </copyright>

namespace PlatformSpellCheck.Tests
{
    using System;
    using System.Linq;

    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Defines test class SpellCheckerTests.
    /// </summary>
    [TestClass]
    public class SpellCheckerTests
    {
        /// <summary>
        /// Defines the test method CheckTest.
        /// </summary>
        [TestMethod]
        public void CheckTest()
        {
            using var spell = new SpellChecker("en-us");
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

        /// <summary>
        /// Defines the test method DoubleDisposeTest.
        /// </summary>
        [TestMethod]
        public void DoubleDisposeTest()
        {
            var spell = new SpellChecker("en-AU");

            spell.Dispose();
            spell.Dispose(); // No error
        }

        /// <summary>
        /// Defines the test method IgnoreTest.
        /// </summary>
        [TestMethod]
        public void IgnoreTest()
        {
            // ReSharper disable once StringLiteralTypo
            // ReSharper disable IdentifierTypo
            const string codez = "codez";
            using var spell = new SpellChecker();
            var examples = spell.Check(codez);

            Assert.IsTrue(examples.Any());

            spell.Ignore(codez);

            examples = spell.Check(codez);

            Assert.AreEqual(examples.Count(), 0);

            // ReSharper restore IdentifierTypo
        }

        /// <summary>
        /// Defines the test method IsLanguageSupportedTest.
        /// </summary>
        [TestMethod]
        public void IsLanguageSupportedTest()
        {
            Assert.IsTrue(SpellChecker.IsLanguageSupported("en-us"));
            Assert.IsFalse(SpellChecker.IsLanguageSupported("zz-zz"));
        }

        /// <summary>
        /// Defines the test method IsPlatformSupportedTest.
        /// </summary>
        [TestMethod]
        public void IsPlatformSupportedTest() => Assert.IsTrue(SpellChecker.IsPlatformSupported());

        /// <summary>
        /// Defines the test method LanguageIdTest.
        /// </summary>
        [TestMethod]
        public void LanguageIdTest()
        {
            using var spell = new SpellChecker("en-AU");
            Assert.AreEqual(spell.LanguageTag, "en-AU");
        }

        /// <summary>
        /// Defines the test method MultiSessionTest.
        /// </summary>
        [TestMethod]
        public void MultiSessionTest()
        {
            var spell1 = new SpellChecker("en-us");
            using var spell2 = new SpellChecker("fr-fr");
            var examples1 = spell1.Suggestions("doog").ToList();
            spell1.Dispose();
            var examples2 = spell2.Suggestions("doog").ToList();

            Assert.AreEqual(examples1.Count(), 10);
            Assert.AreEqual(examples2.Count(), 2);
        }

        /// <summary>
        /// Defines the test method SuggestionsTest.
        /// </summary>
        [TestMethod]
        public void SuggestionsTest()
        {
            using var spell = new SpellChecker();
            var examples = spell.Suggestions("manle");

            Assert.IsTrue(examples.Any());
        }

        /// <summary>
        /// Defines the test method SupportedLanguagesTest.
        /// </summary>
        [TestMethod]
        public void SupportedLanguagesTest() => Assert.IsTrue(SpellChecker.SupportedLanguages.Any());

        /// <summary>
        /// Defines the test method UnsupportedLangTest.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void UnsupportedLangTest()
        {
            using (var spell = new SpellChecker("zz-zz"))
            {
            }
        }
    }
}
