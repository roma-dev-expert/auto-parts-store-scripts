using Microsoft.Extensions.Configuration;
using Serilog;

namespace ExcelTransformer
{
    class Program
    {
        static void Main()
        {
            Log.Logger = new LoggerConfiguration().WriteTo.Console().WriteTo.File("log.txt").CreateLogger();

            IConfigurationRoot configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json")
                .Build(); 

            var excelParser = new Models.ExcelTransformer(configuration.GetSection("Headers"));

            string inputFilePath = FindFirstXlsFile(Directory.GetCurrentDirectory());
            string outputFilePath = "output.xls";

            if (string.IsNullOrEmpty(inputFilePath))
            {
                Log.Error("Не найдены файлы с расширением .xls в текущей директории.");
                return;
            }

            try
            {
                excelParser.TransformToNomenclature(inputFilePath, outputFilePath);
                Log.Information("Преобразование завершено.");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Произошла ошибка.");
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
