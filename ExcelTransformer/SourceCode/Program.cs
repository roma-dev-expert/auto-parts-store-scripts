using Microsoft.Extensions.Configuration;
using Serilog;

namespace ExcelTransformer
{
    class Program
    {
        static void Main()
        {
            Log.Logger = new LoggerConfiguration().WriteTo.Console().WriteTo.File("log.txt").CreateLogger();

            string settingsPath = Path.Combine(Directory.GetCurrentDirectory(), "..", "Settings", "appsettings.json");

            IConfigurationRoot configuration = new ConfigurationBuilder()
                .AddJsonFile(settingsPath)
                .Build(); 

            var excelParser = new Models.ExcelTransformer(configuration);

            string unloadingFrom1CFilePath = FindFirstXlsFile("../Input");
            string nomenclatureFilePath = "../Output/Nomenclature.xls";
            string adsFilePath = "../Output/Ads.xls";

            if (string.IsNullOrEmpty(unloadingFrom1CFilePath))
            {
                Log.Error("No .xls files found in current directory.");
                return;
            }

            try
            {
                excelParser.TransformToNomenclature(unloadingFrom1CFilePath, nomenclatureFilePath);
                excelParser.TransformToAds(nomenclatureFilePath, adsFilePath);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "An error occurred.");
            }
            finally
            {
                Log.CloseAndFlush();
            }
        }

        static string FindFirstXlsFile(string directory)
        {
            string[] xlsFiles = Directory.GetFiles(directory, "*.xls");

            return xlsFiles.FirstOrDefault() ?? "";
        }
    }
}
