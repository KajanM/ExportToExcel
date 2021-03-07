// ExportToExcel/ExportToExcel.Services/Utils/ExcelUtils.cs

using System;

namespace ExportToExcel.Services.Utils
{
    public static class ExcelUtils
    {
        /// <summary>
        /// Works like below
        /// <code>
        /// 1 -> A
        /// 26 -> Z
        /// 28 -> AB
        /// 53 -> BA
        /// </code>
        /// </summary>
        /// <param name="columnNumber">first column starts at one</param>
        public static string GetExcelColumnName(int columnNumber)
        {
            var dividend = columnNumber;
            var columnName = string.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = $"{Convert.ToChar(65 + modulo)}{columnName}";

                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
    }
}