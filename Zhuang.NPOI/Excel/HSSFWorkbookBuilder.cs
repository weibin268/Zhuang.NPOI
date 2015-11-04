using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace Zhuang.NPOI.Excel
{
    /// <summary>
    /// Excel2003
    /// </summary>
    public class HSSFWorkbookBuilder:WorkbookBuilder
    {
        protected override IWorkbook CreateWorkbook()
        {
            return new HSSFWorkbook();
        }
    }
}
