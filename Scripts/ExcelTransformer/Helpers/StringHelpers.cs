using System.Text.RegularExpressions;

namespace ExcelTransformer.Helpers
{
    class StringHelpers
    {
        public static bool IsRussianText(string word)
        {
            return Regex.IsMatch(word, @"^[\p{IsCyrillic}]+$");
        }
        public static bool IsOpeningBracket(string word)
        {
            return word.StartsWith("(");
        }

        public static bool IsClosingBracket(string word)
        {
            return word.EndsWith(")");
        }

        public static bool IsLHOrRH(string word)
        {
            return word == "LH" || word == "RH";
        }

        public static bool IsYearFormat(string word)
        {
            return Regex.IsMatch(word, @"^\d{2}(-\d{2}|\-)?$") || Regex.IsMatch(word, @"^\d{4}-$") || Regex.IsMatch(word, @"^\d{ 4} (-\d{ 4})?$");
        }

        public static string[] ExtractNomenclatureWithoutArticles(string input, string article, string originalArticle)
        {
            var words = SplitBySpace(input);

            int smallestIndex = int.MaxValue;
            int articleIndex = Array.IndexOf(words, article);
            int originalArticleIndex = Array.IndexOf(words, originalArticle);

            if (articleIndex >= 0) smallestIndex = Math.Min(smallestIndex, articleIndex);
            if (originalArticleIndex >= 0) smallestIndex = Math.Min(smallestIndex, originalArticleIndex);

            if (smallestIndex < words.Length) return words.Take(smallestIndex).ToArray();
            
            return new string[0];
        }

        public static bool IsArticle(string word, string article, string originalArticle)
        {
            return word.ToUpper() == article.ToUpper() || word.ToUpper() == originalArticle.ToUpper();
        }

        public static string[] SplitBySpace(string input)
        {
            return input.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        }
    }
}
