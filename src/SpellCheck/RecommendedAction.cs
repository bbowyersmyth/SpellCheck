
namespace PlatformSpellCheck
{
    /// <summary>
    /// Identifies the type of corrective action to be taken for a spelling error
    /// </summary>
    public enum RecommendedAction
    {
        /// <summary>
        /// There are no errors
        /// </summary>
        None = 0,

        /// <summary>
        /// The user should be prompted with a list of suggestions as returned by Suggestions
        /// </summary>
        GetSuggestions = 1,

        /// <summary>
        /// Replace the indicated erroneous text with the text provided in the suggestion. The user does not need to be prompted.
        /// </summary>
        Replace = 2,

        /// <summary>
        /// The user should be prompted to delete the indicated erroneous text
        /// </summary>
        Delete = 3
    }
}
