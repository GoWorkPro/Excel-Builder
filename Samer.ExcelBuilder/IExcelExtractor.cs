using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoWorkPro.ExcelBuilder
{
    public interface IExcelExtractor: IDisposable
    {
        MemoryStream GetStream();
        IDisposable SaveAsFile(string filePath);
        Worksheet GetFirstWorksheet();
        Worksheet GetWorksheet(string name);
        XLWorkbook Workbook { get; }
        Worksheet[] Worksheets { get; }
    }
}
