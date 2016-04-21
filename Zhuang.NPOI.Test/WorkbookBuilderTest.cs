using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Zhuang.NPOI.Excel;
using System.Data;
using Zhuang.Data;
using System.Collections.Generic;

namespace Zhuang.NPOI.Test
{
    /// <summary>
    /// Excel导出测试
    /// </summary>
    [TestClass]
    public class WorkbookBuilderTest
    {
        DbAccessor _dba = DbAccessor.Get();

        /// <summary>
        /// Excel 2003
        /// </summary>
        [TestMethod]
        public void HSSFWorkbookBuilderTest()
        {
            HSSFWorkbookBuilder builder = new HSSFWorkbookBuilder();
            DataTable dt = _dba.QueryDataTable("select top 10 * from sys_product");
            builder.SetDataTable(dt)
                .SetShowColumns(new List<string>() { "ProductName", "ProductCode", "RecordStatus"})
                .SetColumnCaption("产品名称", "ProductName").SetColumnWidth(20)
                .AddOnHeadRowCellCreated((c) =>
                {
                    var style = c.Workbook.CreateCellStyle();
                    var font = c.Workbook.CreateFont();
                    style.SetFont(font);
                    c.Cell.CellStyle = style;

                    font.Boldweight = 600;
                    font.FontHeightInPoints = 11;
                })
                .Build().SaveAs(@"C:\npoitest.xls");
        }


        /// <summary>
        /// Excel 2007
        /// </summary>
        [TestMethod]
        public void XSSFWorkbookBuilderTest()
        {
            XSSFWorkbookBuilder builder = new XSSFWorkbookBuilder();
            DataTable dt = _dba.QueryDataTable("select * from sys_product");
            builder.SetDataTable(dt)
                .SetColumnCaption("产品名称", "ProductName").SetColumnWidth(20)
                .Build().SaveAs(@"C:\npoitest.xlsx");
        }

        
    }
}
