using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;

namespace PlatformSpellCheck
{
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
        /// <param name="lang">A BCP47 language tag that identifies the language for the requested spell checker</param>
        public SpellChecker(string lang)
        {
            var factory = new MsSpellCheckLib.SpellCheckerFactory();

            try
            {
                _spellChecker = factory.CreateSpellChecker(lang);
            }
            finally
            {
                Marshal.ReleaseComObject(factory);
            }
        }

        public static bool IsPlatformSupported()
        {
            return Environment.OSVersion.Version > new Version(6, 2);
        }

        /// <summary>
        /// Determines if the specified language is supported by this spell checker
        /// </summary>
        /// <param name="lang">A BCP47 language tag that identifies the language for the requested spell checker</param>
        /// <returns>true if supported, false otherwise</returns>
        public static bool IsLanguageSupported(string lang)
        {
            var factory = new MsSpellCheckLib.SpellCheckerFactory();

            try
            {
                return (factory.IsSupported(lang) != 0);
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
        /// <param name="word">The text to check</param>
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

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

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
