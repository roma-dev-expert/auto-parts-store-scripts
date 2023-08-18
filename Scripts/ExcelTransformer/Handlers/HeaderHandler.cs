using Microsoft.Extensions.Configuration;

namespace ExcelTransformer.Handlers
{
    public class HeaderHandler
    {
        private readonly IConfigurationSection Headers;

        public HeaderHandler(IConfigurationSection headers)
        {
            Headers = headers;
        }

        public IList<string> ExtractInputHeaders()
        {
            List<string> headers = new List<string>();

            foreach (IConfigurationSection subSection in Headers.GetChildren())
            {
                headers.Add(subSection.Key);
            }

            return headers;
        }

        public IList<string> ExtractOutputHeaders()
        {
            List<string> headers = new List<string>();

            foreach (IConfigurationSection subSection in Headers.GetChildren())
            {
                if (!subSection.GetChildren().Any())
                {
                    headers.Add(subSection.Key);
                    continue;
                }

                foreach (IConfigurationSection subSubSection in subSection.GetChildren())
                {
                    headers.Add(subSubSection.Key);
                }
            }

            return headers;
        }
    }
}
