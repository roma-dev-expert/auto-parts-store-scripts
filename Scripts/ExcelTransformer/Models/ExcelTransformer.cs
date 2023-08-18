using ExcelTransformer.Handlers;
using Microsoft.Extensions.Configuration;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Serilog;

namespace ExcelTransformer.Models
{
    public class ExcelTransformer
    {
        private IConfigurationSection ParserConfiguration;
        private IDictionary<string, StringParser> StringParsers;
        private List<string> InputHeaders;
        private List<string> OutputHeaders;

        public ExcelTransformer(IConfigurationSection parserConfiguration)
        {
            ParserConfiguration = parserConfiguration ?? throw new ArgumentNullException(nameof(parserConfiguration));
            var headerHandler = new HeaderHandler(parserConfiguration);
            InputHeaders = headerHandler.ExtractInputHeaders().ToList();
            OutputHeaders = headerHandler.ExtractOutputHeaders().ToList();
            StringParsers = GetStringParsers();
        }

        public void TransformToNomenclature(string inputPath, string outputPath)
        {
            if (string.IsNullOrEmpty(inputPath) || string.IsNullOrEmpty(outputPath))
            {
                Log.Error("Input or output file path is not provided.");
                return;
            }

            if (!File.Exists(inputPath))
            {
                Log.Error("Input file does not exist.");
                return;
            }

            try
            {
                using (var fs = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
                {
                    var sheet = GetInputSheet(fs);

                    HSSFWorkbook outputWorkbook = new HSSFWorkbook();
                    ISheet outputSheet = outputWorkbook.CreateSheet("OutputSheet");

                    CreateOutputFileHeaders(outputSheet);

                    var columnIndexMap = CreateMapFromInputHeaders(sheet.GetRow(0));

                    int rowIndex = 1;
                    while (sheet.GetRow(rowIndex) != null)
                    {
                        IRow row = sheet.GetRow(rowIndex);
                        IRow newRow = outputSheet.CreateRow(rowIndex);

                        foreach (var header in InputHeaders)
                        {
                            int columnIndex = columnIndexMap[header];
                            string cellValue = GetCellValue(row.GetCell(columnIndex));
                            if (StringParsers.TryGetValue(header, out var stringParser))
                            {
                                var cellValueDictionary = stringParser.Parse(cellValue);
                                foreach (var kvo in cellValueDictionary)
                                {
                                    newRow.CreateCell(OutputHeaders.IndexOf(kvo.Key)).SetCellValue(kvo.Value);
                                }
                            }
                            else newRow.CreateCell(OutputHeaders.IndexOf(header)).SetCellValue(cellValue);
                        }

                        rowIndex++;
                    }

                    using (var fsOutput = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        outputWorkbook.Write(fsOutput);
                    }
                }
                Log.Information("Transformation completed successfully.");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "An error occurred during transformation.");
            }
        }

        public void TransformToAds(string inputPath, string outputPath)
        {

        }

        private IDictionary<string, StringParser> GetStringParsers() 
        {
            var result = new Dictionary<string, StringParser>();
            foreach (var header in InputHeaders)
            {
                if (ParserConfiguration.GetSection(header).GetChildren().Any())
                {
                    var stringParser = new StringParser(ParserConfiguration.GetSection(header));
                    result.Add(header, stringParser);
                }
            }
            return result;
        }

        private ISheet GetInputSheet(FileStream fs)
        {
            HSSFWorkbook workbook = new HSSFWorkbook(fs);
            return workbook.GetSheetAt(0);
        }

        private void CreateOutputFileHeaders(ISheet outputSheet)
        {
            IRow headerRow = outputSheet.CreateRow(0);
            for (int i = 0; i < OutputHeaders.Count; i++)
            {
                headerRow.CreateCell(i).SetCellValue(OutputHeaders[i]);
            }
        }

        private Dictionary<string, int> CreateMapFromInputHeaders(IRow inputHeaderRow)
        {
            Dictionary<string, int> columnIndexMap = new Dictionary<string, int>();

            for (int i = 0; i < InputHeaders.Count; i++)
            {
                string header = InputHeaders[i];
                int columnIndex = -1;

                for (int j = 0; j < inputHeaderRow.LastCellNum; j++)
                {
                    if (inputHeaderRow.GetCell(j)?.ToString() == header)
                    {
                        columnIndex = j;
                        break;
                    }
                }

                columnIndexMap[header] = columnIndex;
            }

            return columnIndexMap;
        }

        private string GetCellValue(ICell cell)
        {
            return cell?.ToString() ?? "";
        }
    }
}
