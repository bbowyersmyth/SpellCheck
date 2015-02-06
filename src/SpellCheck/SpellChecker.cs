using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;

namespace PlatformSpellCheck
{
    /// <summary>
    /// The Spell Checking API permits developers to consume spell checker capability to check text, and get suggestions
    /// </summary>
    public class SpellChecker : IDisposable
    {
        private MsSpellCheckLib.ISpellChecker _spellChecker;

        /// <summary>
        /// Creates a spell checker that supports the current UI language
        /// </summary>
        public SpellChecker()
        {
            var factory = new MsSpellCheckLib.SpellCheckerFactory();

            try
            {
                _spellChecker = factory.CreateSpellChecker(CultureInfo.CurrentUICulture.Name);
            }
            finally
            {
                Marshal.ReleaseComObject(factory);
            }
        }

        /// <summary>
        /// Creates a spell checker that supports the specified language
        /// </summary>
        /// <remarks><see cref="SpellChecker.IsLanguageSupported(System.String)"/> can be called to determine if languageTag is supported</remarks>
        /// <exception cref="System.ArgumentException">languageTag is an empty string, or there is no spell checker available for languageTag</exception>
        /// <param name="languageTag">A BCP47 language tag that identifies the language for the requested spell checker</param>
        public SpellChecker(string languageTag)
        {
            var factory = new MsSpellCheckLib.SpellCheckerFactory();

            try
            {
                _spellChecker = factory.CreateSpellChecker(languageTag);
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
        public static bool IsPlatformSupported()
        {
            return Environment.OSVersion.Version > new Version(6, 2);
        }

        /// <summary>
        /// Determines if the specified language is supported by a registered spell checker
        /// </summary>
        /// <param name="languageTag">A BCP47 language tag that identifies the language for the requested spell checker</param>
        /// <returns>true if supported, false otherwise</returns>
        public static bool IsLanguageSupported(string languageTag)
        {
            var factory = new MsSpellCheckLib.SpellCheckerFactory();

            try
            {
                return (factory.IsSupported(languageTag) != 0);
            }
            finally
            {
                Marshal.ReleaseComObject(factory);
            }
        }

        /// <summary>
        /// Gets the set of BCP47 language tags supported by the spell checker
        /// </summary>
        public static IEnumerable<string> SupportedLanguages
        {
            get
            {
                var factory = new MsSpellCheckLib.SpellCheckerFactory();
                MsSpellCheckLib.IEnumString langs = null;

                try
                {
                    langs = factory.SupportedLanguages;

                    string currentLang;
                    uint fetched;

                    langs.RemoteNext(1, out currentLang, out fetched);

                    while (currentLang != null)
                    {
                        yield return currentLang;
                        langs.RemoteNext(1, out currentLang, out fetched);
                    }
                }
                finally
                {
                    if (langs != null)
                        Marshal.ReleaseComObject(langs);

                    Marshal.ReleaseComObject(factory);
                }
            }
        }

        /// <summary>
        /// Gets the BCP47 language tag this instance of the spell checker supports
        /// </summary>
        public string LanguageTag
        {
            get
            {
                return _spellChecker.languageTag;
            }
        }

        /// <summary>
        /// Treats the provided word as though it were part of the original dictionary.
        /// The word will no longer be considered misspelled, and will also be considered as a candidate for suggestions.
        /// </summary>
        /// <param name="word"></param>
        public void Add(string word)
        {
            _spellChecker.Add(word);
        }

        /// <summary>
        /// Causes occurrences of one word to be replaced by another
        /// </summary>
        /// <param name="from">The incorrectly spelled word to be autocorrected</param>
        /// <param name="to">The correctly spelled word that should replace from</param>
        public void AutoCorrect(string from, string to)
        {
            _spellChecker.AutoCorrect(from, to);
        }

        /// <summary>
        /// Retrieves spelling suggestions for the supplied text
        /// </summary>
        /// <param name="word">The word or phrase to get suggestions for</param>
        /// <returns>The list of suggestions</returns>
        public IEnumerable<string> Suggestions(string word)
        {
            var suggestions = _spellChecker.Suggest(word);

            try
            {
                string currentSuggestion;
                uint fetched;

                suggestions.RemoteNext(1, out currentSuggestion, out fetched);

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
        /// Checks the spelling of the supplied text and returns a collection of spelling errors
        /// </summary>
        /// <param name="text">The text to check</param>
        /// <returns>The results of spell checking</returns>
        public IEnumerable<SpellingError> Check(string text)
        {
            var errors = _spellChecker.Check(text);
            MsSpellCheckLib.ISpellingError currentError = null;

            try
            {
                while ((currentError = errors.Next()) != null)
                {
                    var action = RecommendedAction.None;

                    switch (currentError.CorrectiveAction)
                    {
                        case MsSpellCheckLib.CORRECTIVE_ACTION.CORRECTIVE_ACTION_DELETE:
                            action = RecommendedAction.Delete;
                            break;

                        case MsSpellCheckLib.CORRECTIVE_ACTION.CORRECTIVE_ACTION_GET_SUGGESTIONS:
                            action = RecommendedAction.GetSuggestions;
                            break;

                        case MsSpellCheckLib.CORRECTIVE_ACTION.CORRECTIVE_ACTION_REPLACE:
                            action = RecommendedAction.Replace;
                            break;
                    }

                    yield return new SpellingError()
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
                    Marshal.ReleaseComObject(currentError);

                Marshal.ReleaseComObject(errors);
            }
        }

        /// <summary>
        /// Ignores the provided word for the rest of this session
        /// </summary>
        /// <param name="word">The word to ignore</param>
        public void Ignore(string word)
        {
            _spellChecker.Ignore(word);
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
        /// Disposes resources used by SpellChecker
        /// </summary>
        ~SpellChecker()
        {
            this.Dispose(false);
        }

        private void Dispose(bool disposing)
        {
            if (_spellChecker != null)
            {
                Marshal.ReleaseComObject(_spellChecker);
                _spellChecker = null;
            }
        }
    }
}
