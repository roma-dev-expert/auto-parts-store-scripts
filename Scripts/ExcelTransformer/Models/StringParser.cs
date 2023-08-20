using Microsoft.Extensions.Configuration;

namespace ExcelTransformer.Models
{
    public class StringParser 
    {
        private readonly IConfigurationSection ParserConfiguration;
        private readonly IDictionary<string, IConfigurationSection> ColumnToParseStringConfigs;
        private readonly IDictionary<string, IDictionary<string, IList<string>>> ColumnToSearchKeysDictionary;

        public StringParser(IConfigurationSection parserConfiguration)
        {
            ParserConfiguration = parserConfiguration;
            ColumnToParseStringConfigs = ParserConfiguration.GetChildren().ToList().ToDictionary(cc => cc.Key);
            ColumnToSearchKeysDictionary = GetColumnToSearchKeyDictionary();
        }

        public IDictionary<string, string> Parse(string inputString)
        {
            var result = new Dictionary<string, string>();
            

            foreach (var columnToSearchKeys in ColumnToSearchKeysDictionary)
            {
                if (!columnToSearchKeys.Value.Any()) {
                    result.Add(columnToSearchKeys.Key, inputString);
                    continue;
                }

                var value = "";
                foreach (var searchKeys in columnToSearchKeys.Value)
                {
                    var matchLength = 0;
                    var bestMatchLength = BestMatch(inputString, searchKeys.Value);
                    if (bestMatchLength > matchLength) 
                    {
                        value = searchKeys.Key;
                        matchLength = bestMatchLength;
                    }
                }

                result.Add(columnToSearchKeys.Key, value);

            }
            return result;
        }

        private int BestMatch(string inputString, IList<string> searchKeys)
        {
            int bestMatchLength = 0;

            foreach (string key in searchKeys)
            {
                int startIndex = inputString.IndexOf(key, StringComparison.OrdinalIgnoreCase);
                if (startIndex != -1)
                {
                    int endIndex = startIndex + key.Length;
                    string match = inputString.Substring(endIndex).Trim();
                    if (match.Length > bestMatchLength) bestMatchLength = match.Length;
                }
            }

            return bestMatchLength;
        }

        private IDictionary<string, IDictionary<string, IList<string>>> GetColumnToSearchKeyDictionary()
        {
            var result = new Dictionary<string, IDictionary<string, IList<string>>>();
            foreach (var columnToParseStringConfig in ColumnToParseStringConfigs)
            {
                var searchKeys = new Dictionary<string, IList<string>>();
                foreach (var parseStringConfig in columnToParseStringConfig.Value.GetChildren())
                {
                    searchKeys.Add(parseStringConfig.Key, parseStringConfig.Get<List<string>>() ?? new List<string>());
                }
                result.Add(columnToParseStringConfig.Key, searchKeys);
            }

            return result;
        }
    }
}
