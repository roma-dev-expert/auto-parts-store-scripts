using Microsoft.Extensions.Configuration;
namespace ExcelTransformer.Data
{
    public class CarBrandDatabase
    {
        private readonly IConfiguration Configuration;

        public CarBrandDatabase(IConfiguration configuration)
        {
            Configuration = configuration.GetSection("CarBrandIdentification");
        }

        public string FindBrandByModel(string modelName)
        {
            foreach (var brand in Configuration.GetChildren())
            {
                foreach (var model in brand.GetChildren())
                {
                    if (modelName.IndexOf(model.Value, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return brand.Key;
                    }
                }
            }
            return "";
        }
    }
}
