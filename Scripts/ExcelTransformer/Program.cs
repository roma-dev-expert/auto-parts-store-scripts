using ExcelTransformer.Data;
using ExcelTransformer.Models;
using Microsoft.Extensions.Configuration;
using Serilog;

namespace ExcelTransformer
{
    class Program
    {
        static void Main()
        {
            IConfigurationRoot configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json")
                .Build();


            Log.Logger = new LoggerConfiguration().WriteTo.Console().WriteTo.File("log.txt").CreateLogger();

            var carBrandDatabase = new CarBrandDatabase(configuration);
            string[] headers = configuration.GetSection("Headers").Get<string[]>() ?? new string[0];

            if (headers == null || headers.Length == 0)
            {
                Log.Error("Ошибка: Не удалось получить заголовки из конфигурации.");
                return;
            }

            var excelParser = new ExcelParser(carBrandDatabase, headers);

            string inputFilePath = "input.xls";
            string outputFilePath = "output.xls";

            try
            {
                excelParser.TransformExcel(inputFilePath, outputFilePath);
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
    }
}
