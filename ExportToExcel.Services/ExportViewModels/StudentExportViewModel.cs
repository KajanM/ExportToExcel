// ExportToExcel/ExportToExcel.Services/ExportViewModels/StudentExportViewModel.cs

using System;
using OfficeOpenXml.Attributes;

namespace ExportToExcel.Services.ExportViewModels
{
    public enum Gender
    {
        Male,
        Female
    }
    
    [EpplusTable]
    public class StudentExportViewModel
    {
        
        // [EpplusTableColumn]
        [EpplusIgnore]
        public string Id { get; set; }

        public string Name { get; set; }

        public DateTime DateOfBirth { get; set; }

        public Gender Gender { get; set; }
    }
}