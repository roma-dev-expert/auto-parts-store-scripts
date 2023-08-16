﻿using ExcelTransformer.Data;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ExcelTransformer.Models
{
    public class ExcelParser
    {
        private CarBrandDatabase CarBrandDatabase;
        private string[] Headers;
        public ExcelParser(CarBrandDatabase carBrandDatabase, string[] headers)
        {
            CarBrandDatabase = carBrandDatabase;
            Headers = headers;
        }

        public void TransformExcel(string inputPath, string outputPath)
        {
            using (var fs = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            {
                HSSFWorkbook workbook = new HSSFWorkbook(fs);
                ISheet sheet = workbook.GetSheetAt(0);

                HSSFWorkbook outputWorkbook = new HSSFWorkbook();
                ISheet outputSheet = outputWorkbook.CreateSheet("OutputSheet");

                IRow headerRow = outputSheet.CreateRow(0);
                for (int i = 0; i < Headers.Length; i++)
                {
                    headerRow.CreateCell(i).SetCellValue(Headers[i]);
                }

                int rowIndex = 1; 
                List<string> emptyBrandList = new List<string>();
                while (sheet.GetRow(rowIndex) != null)
                {
                    IRow row = sheet.GetRow(rowIndex);
                    string article = row.GetCell(0).ToString() ?? "";
                    string originalArticle = row.GetCell(1).ToString() ?? "";
                    string nomenclature = row.GetCell(2).ToString() ?? "";
                    string retail = row.GetCell(3).ToString() ?? "";
                    string centralAlmaty = row.GetCell(4).ToString() ?? "";

                    var nomenclatureParser = new NomenclatureParser(nomenclature, article, originalArticle);
                    var parsedNomenclature = nomenclatureParser.ParseNomenclature();
                    var brand = CarBrandDatabase.FindBrandByModel(parsedNomenclature.Model);

                    if (string.IsNullOrEmpty(brand) && brand != "MD" && brand != "AD") emptyBrandList.Add(parsedNomenclature.Model);

                    IRow newRow = outputSheet.CreateRow(rowIndex);
                    newRow.CreateCell(0).SetCellValue(nomenclature);
                    newRow.CreateCell(1).SetCellValue(parsedNomenclature.Group);
                    newRow.CreateCell(2).SetCellValue(parsedNomenclature.Addition);
                    newRow.CreateCell(3).SetCellValue(parsedNomenclature.Subgroup);
                    newRow.CreateCell(4).SetCellValue(brand);
                    newRow.CreateCell(5).SetCellValue(parsedNomenclature.Model);
                    newRow.CreateCell(6).SetCellValue(parsedNomenclature.Year);
                    newRow.CreateCell(7).SetCellValue(article);
                    newRow.CreateCell(8).SetCellValue(originalArticle);
                    newRow.CreateCell(9).SetCellValue(retail);
                    newRow.CreateCell(10).SetCellValue(centralAlmaty);

                    rowIndex++;
                }

                emptyBrandList = emptyBrandList.Distinct().ToList();

                using (var fsOutput = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    outputWorkbook.Write(fsOutput);
                }
            }
        }
    }
}
