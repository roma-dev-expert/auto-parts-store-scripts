using ExcelTransformer.Helpers;
using Serilog;
using System.Text.RegularExpressions;

namespace ExcelTransformer.Models
{
    public class NomenclatureParser
    {
        private string[] Words;
        private string Nomenclature;
        private ParsedNomenclature ParsedNomenclature;

        public NomenclatureParser(string nomenclature, string article, string originalArticle)
        {
            Nomenclature = nomenclature;
            Words = StringHelpers.ExtractNomenclatureWithoutArticles(nomenclature, article, originalArticle);
            ParsedNomenclature = new ParsedNomenclature();
        }

        public ParsedNomenclature ParseNomenclature()
        {
            if (Words.Length == 0)
            {
                Log.Warning($"No words found for parsing {Nomenclature}");
                return ParsedNomenclature;
            }

            ParsedNomenclature.Addition = ExtractAddition();
            ParsedNomenclature.Subgroup = ExtractSubgroup();
            ParsedNomenclature.Group = ExtractGroup();
            ParsedNomenclature.Model = ExtractModel();
            ParsedNomenclature.Year = ExtractYear();

            return ParsedNomenclature;
        }

        private string ExtractAddition()
        {
            string combinedText = string.Join(" ", Words);
            Match match = Regex.Match(combinedText, @"\(([^)]*)\)");
            return match.Success ? "(" + match.Groups[1].Value.Trim() + ")" : "";
        }

        private string ExtractSubgroup()
        {
            string combinedText = string.Join(" ", Words);
            return combinedText.Contains(" LH ") ? "левая" : combinedText.Contains(" RH ") ? "правая" : "";
        }

        private string ExtractGroup()
        {
            int currentIndex = 0;
            string group = "";

            while (currentIndex == 0 || (currentIndex < Words.Length && !StringHelpers.IsOpeningBracket(Words[currentIndex]) &&
                   StringHelpers.IsRussianText(Words[currentIndex]) && !StringHelpers.IsLHOrRH(Words[currentIndex])))
            {
                group += Words[currentIndex] + " ";
                currentIndex++;
            }

            return group.Trim();
        }

        private string ExtractModel()
        {
            int currentIndex = GetCurrentIndex();
            string model = "";

            while (currentIndex < Words.Length && !StringHelpers.IsYearFormat(Words[currentIndex])
                && !StringHelpers.IsRussianText(Words[currentIndex]))
            {
                model += Words[currentIndex] + " ";
                currentIndex++;
            }

            return model.Trim();
        }

        private string ExtractYear()
        {
            int currentIndex = GetCurrentIndex();
            string year = "";

            while (currentIndex < Words.Length && StringHelpers.IsYearFormat(Words[currentIndex]))
            {
                year += Words[currentIndex] + " ";
                currentIndex++;
            }

            return year.Trim();
        }

        private int GetCurrentIndex()
        {
            string combinedText = string.Join(" ", ParsedNomenclature.Group, ParsedNomenclature.Subgroup, ParsedNomenclature.Addition, ParsedNomenclature.Model, ParsedNomenclature.Year);
            return StringHelpers.SplitBySpace(combinedText.Trim()).Length;
        }
    }
}
