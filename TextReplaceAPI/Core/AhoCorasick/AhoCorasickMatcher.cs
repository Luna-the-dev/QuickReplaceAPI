using TextReplaceAPI.Data;

namespace TextReplaceAPI.Core.AhoCorasick
{
    internal static class AhoCorasickMatcher
    {
        public static AhoCorasickStringSearcher CreateMatcher(Dictionary<string, string> replacePhrases, bool caseSensitive)
        {
            // construct the automaton and fill it with the phrases to search for
            // also create a list of the replacement phrases to go alongside the 
            var matcher = new AhoCorasickStringSearcher(caseSensitive);
            foreach (var searchWord in replacePhrases)
            {
                matcher.AddItem(searchWord.Key);
            }
            matcher.CreateFailureFunction();

            return matcher;
        }
    }
}
