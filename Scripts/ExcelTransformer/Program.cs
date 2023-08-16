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
            var ExcelParser = new ExcelParser(CarBrandDatabase);

            string inputFilePath = "input.xls";
            string outputFilePath = "output.xls";

            ExcelParser.TransformExcel(inputFilePath, outputFilePath);

            Console.WriteLine("Преобразование завершено.");
        }

        
    }
}