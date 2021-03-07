// ExportToExcel/ExportToExcel.UnitTests/ExportToExcelServiceTests.cs

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using ExportToExcel.Services.Concrete;
using ExportToExcel.Services.Dtos;
using ExportToExcel.Services.ExportViewModels;
using Xunit;

namespace ExportToExcel.UnitTests
{
    public class ExportToExcelServiceTests
    {
        [Fact]
        public async Task Export_Should_Succeed()
        {
            #region Arrange

            var students = new List<StudentExportViewModel>();
            var religions = new List<ReligionExportViewModel>();
            for (var i = 0; i < 5; i++)
            {
                students.Add(new StudentExportViewModel
                {
                    Id = Guid.NewGuid().ToString(),
                    Name = $"student-{i}",
                    DateOfBirth = DateTime.Today,
                    Gender = Gender.Male,
                });

                religions.Add(new ReligionExportViewModel
                {
                    Name = $"religion-{i}"
                });
            }

            var studentsSheetArgs = new WorksheetArgs<StudentExportViewModel>
            {
                Records = students,
                Title = "Students"
            };
            var religionsSheetArgs = new WorksheetArgs<ReligionExportViewModel>
            {
                Records = religions,
                Title = "Religions"
            };

            #endregion

            #region Act

            var excelService = new GenerateExcelService();

            var data = await excelService
                .SetWorkbookTitle("test")
                .SetDateFormat("dd-MM-YYYY")
                .AddWorksheet(studentsSheetArgs)
                .AddWorksheet(religionsSheetArgs)
                .GenerateExcelAsync();

            #endregion

            #region Assert

            Assert.NotEmpty(data);

            // await File.WriteAllBytesAsync("Should_Be_Able_To_Upload_And_Download_That_File.xlsx", data);

            #endregion
        }
    }
}