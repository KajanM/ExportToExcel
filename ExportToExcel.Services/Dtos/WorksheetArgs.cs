// ExportToExcel/ExportToExcel.Services/Dtos/WorksheetArgs.cs

using System.Collections.Generic;

namespace ExportToExcel.Services.Dtos
{
    /// <summary>
    /// Contains arguments required to generate an Excel worksheet
    /// </summary>
    public class WorksheetArgs<ViewModelType>
    {
        public string Title { get; set; }

        public IEnumerable<ViewModelType> Records { get; set; }
    }
}