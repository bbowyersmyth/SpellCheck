// <copyright file="SpellingError.cs" company="Bruce Bowyer-Smyth">
//     Copyright © Bruce Bowyer-Smyth
// </copyright>

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
        /// <value>The start index.</value>
        public long StartIndex { get; internal set; }

        /// <summary>
        /// Gets the length of the erroneous text
        /// </summary>
        /// <value>The length.</value>
        public long Length { get; internal set; }

        /// <summary>
        /// Indicates which corrective action should be taken for the spelling error
        /// </summary>
        /// <value>The recommended action.</value>
        public RecommendedAction RecommendedAction { get; internal set; }

        /// <summary>
        /// Gets the text to use as replacement text when the corrective action is replace
        /// </summary>
        /// <value>The recommended replacement.</value>
        public string RecommendedReplacement { get; internal set; }
    }
}
