using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Text.SpellCheck
{
    public class SpellingError
    {
        public long StartIndex { get; internal set; }
        public long Length { get; internal set; }
        public RecommendedAction RecommendedAction { get; internal set; }
        public string RecommendedReplacement { get; internal set; }
    }
}
