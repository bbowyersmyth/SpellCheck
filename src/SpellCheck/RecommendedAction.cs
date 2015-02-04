using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Text.SpellCheck
{
    public enum RecommendedAction
    {
        None = 0,
        GetSuggestions = 1,
        Replace = 2,
        Delete = 3
    }
}
