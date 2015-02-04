
namespace PlatformSpellCheck
{
    public class SpellingError
    {
        public long StartIndex { get; internal set; }
        public long Length { get; internal set; }
        public RecommendedAction RecommendedAction { get; internal set; }
        public string RecommendedReplacement { get; internal set; }
    }
}
