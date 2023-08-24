using Microsoft.Extensions.Configuration;

namespace ExcelTransformer.Handlers
{
    public static class HeaderHandler
    {
        public static IList<string> ExtractHeaders(IConfigurationSection headers)
        {
            List<string> result = new List<string>();

            foreach (IConfigurationSection subSection in headers.GetChildren())
            {
                result.Add(subSection.Key);
            }

            return result;
        }

        public static IList<string> ExtractNestedHeaders(IConfigurationSection headers)
        {
            List<string> result = new List<string>();

            foreach (IConfigurationSection subSection in headers.GetChildren())
            {
                if (!subSection.GetChildren().Any())
                {
                    result.Add(subSection.Key);
                    continue;
                }

                foreach (IConfigurationSection subSubSection in subSection.GetChildren())
                {
                    result.Add(subSubSection.Key);
                }
            }

            return result;
        }
    }
}
