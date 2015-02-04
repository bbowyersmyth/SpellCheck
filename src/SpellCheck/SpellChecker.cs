using System;
using System.Collections.Generic;
using System.Globalization;

namespace PlatformSpellCheck
{
    public class SpellChecker
    {
        private readonly MsSpellCheckLib.ISpellChecker _spellChecker;

        public SpellChecker()
        {
            var factory = new MsSpellCheckLib.SpellCheckerFactory();
            _spellChecker = factory.CreateSpellChecker(CultureInfo.CurrentUICulture.Name);
        }

        public SpellChecker(string lang)
        {
            var factory = new MsSpellCheckLib.SpellCheckerFactory();
            _spellChecker = factory.CreateSpellChecker(lang);
        }

        public static bool IsPlatformSupported()
        {
            return Environment.OSVersion.Version > new Version(6, 2);
        }

        public static bool IsLanguageSupported(string lang)
        {
            var factory = new MsSpellCheckLib.SpellCheckerFactory();

            return (factory.IsSupported(lang) != 0);
        }

        public static IEnumerable<string> SupportedLanguages
        {
            get
            {
                var factory = new MsSpellCheckLib.SpellCheckerFactory();

                var langs = factory.SupportedLanguages;
                string currentLang;
                uint fetched;

                langs.RemoteNext(1, out currentLang, out fetched);

                while (currentLang != null)
                {
                    yield return currentLang;
                    langs.RemoteNext(1, out currentLang, out fetched);
                }
            }
        }

        public IEnumerable<string> Suggestions(string word)
        {
            var suggestions = _spellChecker.Suggest(word);
            string currentSuggestion;
            uint fetched;

            suggestions.RemoteNext(1, out currentSuggestion, out fetched);

            while (currentSuggestion != null)
            {
                yield return currentSuggestion;
                suggestions.RemoteNext(1, out currentSuggestion, out fetched);
            }
        }

        public IEnumerable<SpellingError> Check(string text)
        {
            var errors = _spellChecker.Check(text);
            MsSpellCheckLib.ISpellingError currentError;

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
            }
        }
    }
}
