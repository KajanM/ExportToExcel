// ExportToExcel/ExportToExcel.Services/IGenerateExcelService.cs

using System.Threading.Tasks;
using ExportToExcel.Services.Dtos;

namespace ExportToExcel.Services
{
    public interface IGenerateExcelService
    {
        IGenerateExcelService SetWorkbookTitle(string title);
    
        IGenerateExcelService AddWorksheet<ViewModelType>(WorksheetArgs<ViewModelType> worksheetArgs);
    
        IGenerateExcelService SetDateFormat(string dateFormat);
    
        Task<byte[]> GenerateExcelAsync();
    }
}