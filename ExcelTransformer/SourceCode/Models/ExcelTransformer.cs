using ExcelTransformer.Handlers;
using Microsoft.Extensions.Configuration;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Serilog;

namespace ExcelTransformer.Models
{
    public class ExcelTransformer
    {
        private IDictionary<string, StringParser> NomenclatureStringParsers;
        private IDictionary<string, StringParser> AdsStringParsers;
        private IDictionary<string, Template> Templates;
        private List<string> InputHeaders;
        private List<string> NomenclatureHeaders;
        private List<string> AdsHeaders;

        public ExcelTransformer(IConfigurationRoot parserConfiguration)
        {
            InputHeaders = HeaderHandler.ExtractHeaders(parserConfiguration.GetSection("NomenclatureHeaders")).ToList();
            NomenclatureHeaders = HeaderHandler.ExtractNestedHeaders(parserConfiguration.GetSection("NomenclatureHeaders")).ToList();
            Templates = LoadTemplates("../Templates");
            AdsHeaders = HeaderHandler.ExtractNestedHeaders(parserConfiguration.GetSection("AdsHeaders")).ToList();
            AdsHeaders.AddRange(Templates.Keys);
            NomenclatureStringParsers = GetStringParsers(parserConfiguration.GetSection("NomenclatureHeaders"), InputHeaders);
            AdsStringParsers = GetStringParsers(parserConfiguration.GetSection("AdsHeaders"), NomenclatureHeaders);
        }

        public void TransformToNomenclature(string inputPath, string outputPath)
        {
            CheckInputOutputPath(inputPath, outputPath);

            try
            {
                using (var fs = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
                {
                    var sheet = GetInputSheet(fs);

                    HSSFWorkbook outputWorkbook = new HSSFWorkbook();
                    ISheet outputSheet = outputWorkbook.CreateSheet("OutputSheet");

                    CreateSheetHeaders(outputSheet, NomenclatureHeaders);

                    var columnIndexMap = CreateMapFromHeaders(sheet.GetRow(0), InputHeaders);

                    int rowIndex = 1;
                    while (sheet.GetRow(rowIndex) != null)
                    {
                        IRow row = sheet.GetRow(rowIndex);
                        IRow newRow = outputSheet.CreateRow(rowIndex);

                        foreach (var header in InputHeaders)
                        {
                            int columnIndex = columnIndexMap[header];
                            string cellValue = GetCellValue(row.GetCell(columnIndex));
                            if (NomenclatureStringParsers.TryGetValue(header, out var stringParser))
                            {
                                var cellValueDictionary = stringParser.Parse(cellValue);
                                foreach (var kvo in cellValueDictionary)
                                {
                                    newRow.CreateCell(NomenclatureHeaders.IndexOf(kvo.Key)).SetCellValue(kvo.Value);
                                }
                            }
                            else newRow.CreateCell(NomenclatureHeaders.IndexOf(header)).SetCellValue(cellValue);
                        }

                        rowIndex++;
                    }

                    using (var fsOutput = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        outputWorkbook.Write(fsOutput);
                    }
                }
                Log.Information("Transformation 1C output file to Nomenclature completed successfully.");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "An error occurred during transformation.");
            }
        }

        public void TransformToAds(string inputPath, string outputPath)
        {
            CheckInputOutputPath(inputPath, outputPath);
            try
            {
                using (var fs = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
                {
                    var sheet = GetInputSheet(fs);

                    HSSFWorkbook outputWorkbook = new HSSFWorkbook();
                    ISheet outputSheet = outputWorkbook.CreateSheet("OutputSheet");

                    CreateSheetHeaders(outputSheet, AdsHeaders);

                    var columnIndexMap = CreateMapFromHeaders(sheet.GetRow(0), NomenclatureHeaders);

                    int rowIndex = 1;
                    while (sheet.GetRow(rowIndex) != null)
                    {
                        IRow row = sheet.GetRow(rowIndex);
                        IRow newRow = outputSheet.CreateRow(rowIndex);

                        foreach (var header in NomenclatureHeaders)
                        {
                            var cellValue = "";


                            if (columnIndexMap.TryGetValue(header, out int columnIndex))
                            {
                                cellValue = GetCellValue(row.GetCell(columnIndex));
                            }

                            if (AdsStringParsers.TryGetValue(header, out var stringParser))
                            {
                                var cellValueDictionary = stringParser.Parse(cellValue);
                                foreach (var kvo in cellValueDictionary)
                                {
                                    newRow.CreateCell(AdsHeaders.IndexOf(kvo.Key)).SetCellValue(kvo.Value);
                                }
                                continue;
                            }

                            if (AdsHeaders.Contains(header))
                            {
                                newRow.CreateCell(AdsHeaders.IndexOf(header)).SetCellValue(cellValue);
                            }
                        }

                        foreach(var header in Templates.Keys) {
                            var transformedText = TransformUsingTemplate(Templates[header], columnIndexMap, row);
                            newRow.CreateCell(AdsHeaders.IndexOf(header)).SetCellValue(transformedText);
                        }
                            
                        rowIndex++;
                    }

                    using (var fsOutput = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        outputWorkbook.Write(fsOutput);
                    }
                }
                Log.Information("Transformation Nomenclature to Ads completed successfully.");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "An error occurred during transformation.");
            }

        }

        private void CheckInputOutputPath(string inputPath, string outputPath)
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
        }

        private IDictionary<string, StringParser> GetStringParsers(IConfigurationSection section, List<string> headers) 
        {
            var result = new Dictionary<string, StringParser>();
            foreach (var header in headers)
            {
                if (section.GetSection(header).GetChildren().Any())
                {
                    var stringParser = new StringParser(section.GetSection(header));
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

        private void CreateSheetHeaders(ISheet sheet, List<string> headers)
        {
            IRow headerRow = sheet.CreateRow(0);
            for (int i = 0; i < headers.Count; i++)
            {
                headerRow.CreateCell(i).SetCellValue(headers[i]);
            }
        }

        private Dictionary<string, int> CreateMapFromHeaders(IRow inputHeaderRow, List<string> headers)
        {
            Dictionary<string, int> columnIndexMap = new Dictionary<string, int>();

            for (int i = 0; i < headers.Count; i++)
            {
                string header = headers[i];
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

        private IDictionary<string, Template>  LoadTemplates(string directory)
        {
            string[] txtFiles = Directory.GetFiles(directory, "*.txt");
            List<Template> templates = new List<Template>();

            foreach (var txtFilePath in txtFiles)
            {
                templates.Add(new Template(txtFilePath));
            }

            return templates.ToDictionary(t => t.Name);
        }

        private string TransformUsingTemplate(Template template, IDictionary<string, int> columnIndexMap, IRow row)
        {
            var transformedText = template.Text;
            foreach (var column in template.Columns)
            {
                int columnIndex = columnIndexMap[column];
                string cellValue = GetCellValue(row.GetCell(columnIndex));
                cellValue = string.IsNullOrEmpty(cellValue) ? "" : cellValue + " "; 
                transformedText = transformedText.Replace($"[{column}]", cellValue);
            }

            return transformedText;
        }
    }
}
