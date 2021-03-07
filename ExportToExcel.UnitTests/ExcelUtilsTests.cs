// ExportToExcel/ExportToExcel.UnitTests/ExcelUtilsTests.cs

using System.Collections.Generic;
using ExportToExcel.Services.Utils;
using Xunit;

namespace ExportToExcel.UnitTests
{
    public class ExcelUtilsTests
    {
        [Fact]
        public void GetExcelColumnName_Should_Behave_As_Expected()
        {
            #region Arrange

            var inputVsExpectedOutput = new Dictionary<int, string>
            {
                {1, "A"},
                {26, "Z"},
                {28, "AB"},
                {53, "BA"}
            };

            #endregion

            #region ActAndAssert

            foreach (var (input, expectedOutput) in inputVsExpectedOutput)
            {
                var output = ExcelUtils.GetExcelColumnName(input);

                Assert.Equal(output, expectedOutput);
            }

            #endregion
        }
    }
}