using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Samer.ExcelBuilder
{
    public interface IExcelBuilder
    {
        IExcelExtractor Build();
    }
}
