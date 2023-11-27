using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoWorkPro.ExcelBuilder
{
    public interface IExcelBuilder
    {
        IExcelExtractor Build();
    }
}
