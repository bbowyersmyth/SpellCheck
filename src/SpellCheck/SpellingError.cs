
namespace PlatformSpellCheck
{
    /// <summary>
    /// Provides information about a spelling error
    /// </summary>
    public class SpellingError
    {
        /// <summary>
        /// Gets the position in the checked text where the error begins
        /// </summary>
        public long StartIndex { get; internal set; }

        /// <summary>
        /// Gets the length of the erroneous text
        /// </summary>
        public long Length { get; internal set; }

        /// <summary>
        /// Indicates which corrective action should be taken for the spelling error
        /// </summary>
        public RecommendedAction RecommendedAction { get; internal set; }

        /// <summary>
        /// Gets the text to use as replacement text when the corrective action is replace
        /// </summary>
        public string RecommendedReplacement { get; internal set; }
    }
}
