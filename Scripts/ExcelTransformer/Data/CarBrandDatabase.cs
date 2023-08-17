using Microsoft.Extensions.Configuration;
using Serilog;
using Serilog.Core;
using System;

namespace ExcelTransformer.Data
{
    public class CarBrandDatabase
    {
        private readonly IConfiguration Configuration;
        private Logger FileLogger;

        public CarBrandDatabase(IConfiguration configuration)
        {
            Configuration = configuration.GetSection("CarBrandIdentification");
            FileLogger = new LoggerConfiguration()
                .WriteTo.File("log.txt")
                .CreateLogger();
        }

        public string FindBrandByModel(string modelName)
        {
            foreach (var brand in Configuration.GetChildren())
            {
                foreach (var model in brand.GetChildren())
                {
                    var modelValue = model.Value ?? "";
                    if (ModelMatchesBrand(modelName, modelValue))
                    {
                        LogMatchedModel(modelName, brand.Key, modelValue);
                        return brand.Key;
                    }
                }
            }

            LogNoBrandFound(modelName);
            return "";
        }

        private bool ModelMatchesBrand(string modelName, string modelValue)
        {
            return modelName.Contains(modelValue, StringComparison.OrdinalIgnoreCase);
        }

        private void LogMatchedModel(string modelName, string brandName, string modelValue)
        {
            FileLogger.Information($"Model '{modelName}' matches brand '{brandName}' for model value '{modelValue}'.");
        }

        private void LogNoBrandFound(string modelName)
        {
            FileLogger.Information($"No brand found for model '{modelName}'.");
        }
    }
}
