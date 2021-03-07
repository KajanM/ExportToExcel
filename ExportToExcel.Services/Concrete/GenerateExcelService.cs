using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using ExportToExcel.Services.Dtos;
using ExportToExcel.Services.Utils;
using OfficeOpenXml;
using OfficeOpenXml.Attributes;
using OfficeOpenXml.Style;

namespace ExportToExcel.Services.Concrete
{
    public class GenerateExcelService : IGenerateExcelService
    {
        private readonly ExcelPackage _excelPackage;
        private string _dateFormat = "dd/MM/YYYY";

        public GenerateExcelService()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _excelPackage = new ExcelPackage();
        }

        public IGenerateExcelService SetWorkbookTitle(string title)
        {
            _excelPackage.Workbook.Properties.Title = title;

            return this;
        }

        public IGenerateExcelService AddWorksheet<ViewModelType>(WorksheetArgs<ViewModelType> worksheetArgs)
        {
            var worksheet = _excelPackage.Workbook.Worksheets.Add(worksheetArgs.Title);

            LoadDataToWorksheet(worksheet, worksheetArgs.Records);
            FormatDateDisplayValue<ViewModelType>(worksheet);

            var headerRange = GetHeaderRange(worksheet);
            StyleHeaderRow(worksheet, headerRange);
            worksheet.Cells[headerRange].AutoFilter = true;

            worksheet.Cells.AutoFitColumns();

            worksheet.IgnoredErrors.Add(worksheet.Cells[$"A1:{worksheet.Dimension.End.Address}"]);
            //
            return this;
        }

        public IGenerateExcelService SetDateFormat(string dateFormat)
        {
            _dateFormat = dateFormat;

            return this;
        }

        public async Task<byte[]> GenerateExcelAsync()
        {
            return await _excelPackage.GetAsByteArrayAsync();
        }

        private static void LoadDataToWorksheet<T>(ExcelWorksheet worksheet, IEnumerable<T> records)
        {
            worksheet.Cells["A1"].LoadFromCollection(records, true);
        }


        private void FormatDateDisplayValue<T>(ExcelWorksheet worksheet)
        {
            var properties = GetNotIgnoredProperties<T>();
            var dateColumnIndexes = new List<int>();

            for (var i = 0; i < properties.Length; i++)
            {
                var currentPropertyType = properties[i].PropertyType;
                if (currentPropertyType == typeof(DateTime) || currentPropertyType == typeof(DateTime?))
                {
                    dateColumnIndexes.Add(i + 1); // first column index starts in 1
                }
            }

            dateColumnIndexes.ForEach(columnIndex =>
            {
                worksheet.Column(columnIndex).Style.Numberformat.Format = _dateFormat;
            });
        }

        private static PropertyInfo[] GetNotIgnoredProperties<ViewModelType>()
        {
            return typeof(ViewModelType)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(p => !Attribute.IsDefined(p, typeof(EpplusTableColumnAttribute)) &&
                            !Attribute.IsDefined(p, typeof(EpplusIgnore)))
                .ToArray();
        }

        private static string GetHeaderRange(ExcelWorksheet worksheet)
        {
            var lastColumnAddress = $"{ExcelUtils.GetExcelColumnName(worksheet.Dimension.End.Column)}1";
        
            return $"A1:{lastColumnAddress}";
        }

        private static void StyleHeaderRow(ExcelWorksheet worksheet, string headerRange)
        {
            using (var rng = worksheet.Cells[headerRange])
            {
                rng.Style.Font.Bold = true;
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(Color.DarkGray);
            }
        }
    }
}