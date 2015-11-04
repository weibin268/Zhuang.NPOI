using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace Zhuang.NPOI.Excel
{
    /// <summary>
    /// Excel2007
    /// </summary>
    public class XSSFWorkbookBuilder:WorkbookBuilder
    {
        protected override IWorkbook CreateWorkbook()
        {
            return new XSSFWorkbook();
        }
    }
}
