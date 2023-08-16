using ExcelTransformer.Data;
using ExcelTransformer.Models;
using Microsoft.Extensions.Configuration;

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

            var CarBrandDatabase = new CarBrandDatabase(configuration);
            string[] headers = configuration.GetSection("Headers").Get<string[]>() ?? new string[0];
            var ExcelParser = new ExcelParser(CarBrandDatabase, headers);

            string inputFilePath = "input.xls";
            string outputFilePath = "output.xls";

            ExcelParser.TransformExcel(inputFilePath, outputFilePath);

            Console.WriteLine("Преобразование завершено.");
        }

        
    }
}