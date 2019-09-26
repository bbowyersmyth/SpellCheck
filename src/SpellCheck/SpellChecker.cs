// <copyright file="SpellChecker.cs" company="Bruce Bowyer-Smyth">
//     Copyright © Bruce Bowyer-Smyth
// </copyright>

namespace PlatformSpellCheck
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Runtime.InteropServices;

    using MsSpellCheckLib;

    /// <summary>
    /// The Spell Checking API permits developers to consume spell checker capability to check text, and get suggestions
    /// Implements the <see cref="System.IDisposable" />
    /// </summary>
    /// <seealso cref="System.IDisposable" />
    public class SpellChecker : IDisposable
    {
        /// <summary>
        /// The spell checker
        /// </summary>
        private ISpellChecker spellChecker;

        /// <summary>
        /// Initializes a new instance of the <see cref="SpellChecker"/> class.
        /// Creates a spell checker that supports the current UI language
        /// </summary>
        public SpellChecker()
        {
            var factory = new SpellCheckerFactory();

            try
            {
                this.spellChecker = factory.CreateSpellChecker(CultureInfo.CurrentUICulture.Name);
            }
            finally
            {
                Marshal.ReleaseComObject(factory);
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SpellChecker"/> class.
        /// Creates a spell checker that supports the specified language
        /// </summary>
        /// <param name="languageTag">
        /// A BCP47 language tag that identifies the language for the requested spell checker
        /// </param>
        /// <exception cref="System.ArgumentException">
        /// languageTag is an empty string, or there is no spell checker available for languageTag
        /// </exception>
        /// <remarks>
        /// <see cref="SpellChecker.IsLanguageSupported(string)"/> can be called to determine if languageTag is supported
        /// </remarks>
        public SpellChecker(string languageTag)
        {
            var factory = new SpellCheckerFactory();

            try
            {
                this.spellChecker = factory.CreateSpellChecker(languageTag);
            }
            finally
            {
                Marshal.ReleaseComObject(factory);
            }
        }

        /// <summary>
        /// Finalizes an instance of the <see cref="SpellChecker"/> class.
        /// </summary>
        ~SpellChecker()
        {
            this.Dispose(false);
        }

        /// <summary>
        /// Gets the set of BCP47 language tags supported by the spell checker
        /// </summary>
        /// <value>The supported languages.</value>
        public static IEnumerable<string> SupportedLanguages
        {
            get
            {
                var factory = new SpellCheckerFactory();
                IEnumString languages = null;

                try
                {
                    languages = factory.SupportedLanguages;

                    languages.RemoteNext(1, out var currentLang, out var fetched);

                    while (currentLang != null)
                    {
                        yield return currentLang;
                        languages.RemoteNext(1, out currentLang, out fetched);
                    }
                }
                finally
                {
                    if (languages != null)
                    {
                        Marshal.ReleaseComObject(languages);
                    }

                    Marshal.ReleaseComObject(factory);
                }
            }
        }

        /// <summary>
        /// Gets the BCP47 language tag this instance of the spell checker supports
        /// </summary>
        /// <value>The language tag.</value>
        public string LanguageTag => this.spellChecker.languageTag;

        /// <summary>
        /// Determines if the specified language is supported by a registered spell checker
        /// </summary>
        /// <param name="languageTag">A BCP47 language tag that identifies the language for the requested spell checker</param>
        /// <returns>true if supported, false otherwise</returns>
        public static bool IsLanguageSupported(string languageTag)
        {
            var factory = new SpellCheckerFactory();

            try
            {
                return factory.IsSupported(languageTag) != 0;
            }
            finally
            {
                Marshal.ReleaseComObject(factory);
            }
        }

        /// <summary>
        /// Determines if the current operating system is supports the Windows Spell Checking API
        /// </summary>
        /// <returns>true if OS is supported, false otherwise</returns>
        public static bool IsPlatformSupported() => Environment.OSVersion.Version > new Version(6, 2);

        /// <summary>
        /// Treats the provided word as though it were part of the original dictionary.
        /// The word will no longer be considered misspelled, and will also be considered as a candidate for suggestions.
        /// </summary>
        /// <param name="word">The word.</param>
        public void Add(string word) => this.spellChecker.Add(word);

        /// <summary>
        /// Causes occurrences of one word to be replaced by another
        /// </summary>
        /// <param name="from">The incorrectly spelled word to be auto-corrected</param>
        /// <param name="to">The correctly spelled word that should replace from</param>
        public void AutoCorrect(string from, string to) => this.spellChecker.AutoCorrect(from, to);

        /// <summary>
        /// Checks the spelling of the supplied text and returns a collection of spelling errors
        /// </summary>
        /// <param name="text">The text to check</param>
        /// <returns>The results of spell checking</returns>
        public IEnumerable<SpellingError> Check(string text)
        {
            var errors = this.spellChecker.Check(text);
            ISpellingError currentError = null;

            try
            {
                while ((currentError = errors.Next()) != null)
                {
                    var action = RecommendedAction.None;

                    switch (currentError.CorrectiveAction)
                    {
                        case CORRECTIVE_ACTION.CORRECTIVE_ACTION_DELETE:
                            action = RecommendedAction.Delete;
                            break;

                        case CORRECTIVE_ACTION.CORRECTIVE_ACTION_GET_SUGGESTIONS:
                            action = RecommendedAction.GetSuggestions;
                            break;

                        case CORRECTIVE_ACTION.CORRECTIVE_ACTION_REPLACE:
                            action = RecommendedAction.Replace;
                            break;
                    }

                    yield return new SpellingError
                                     {
                                         StartIndex = currentError.StartIndex,
                                         Length = currentError.Length,
                                         RecommendedAction = action,
                                         RecommendedReplacement = currentError.Replacement
                                     };

                    Marshal.ReleaseComObject(currentError);
                }
            }
            finally
            {
                if (currentError != null)
                {
                    Marshal.ReleaseComObject(currentError);
                }

                Marshal.ReleaseComObject(errors);
            }
        }

        /// <summary>
        /// Disposes resources used by SpellChecker
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Ignores the provided word for the rest of this session
        /// </summary>
        /// <param name="word">The word to ignore</param>
        public void Ignore(string word) => this.spellChecker.Ignore(word);

        /// <summary>
        /// Retrieves spelling suggestions for the supplied text
        /// </summary>
        /// <param name="word">The word or phrase to get suggestions for</param>
        /// <returns>The list of suggestions</returns>
        public IEnumerable<string> Suggestions(string word)
        {
            var suggestions = this.spellChecker.Suggest(word);

            try
            {
                suggestions.RemoteNext(1, out var currentSuggestion, out var fetched);

                while (currentSuggestion != null)
                {
                    yield return currentSuggestion;
                    suggestions.RemoteNext(1, out currentSuggestion, out fetched);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(suggestions);
            }
        }

        /// <summary>
        /// Releases unmanaged and - optionally - managed resources.
        /// </summary>
        /// <param name="disposing"><c>true</c> to release both managed and unmanaged resources; <c>false</c> to release only unmanaged resources.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (this.spellChecker != null)
            {
                Marshal.ReleaseComObject(this.spellChecker);
                this.spellChecker = null;
            }
        }
    }
}