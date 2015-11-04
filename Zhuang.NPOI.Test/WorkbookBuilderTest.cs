using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Zhuang.NPOI.Excel;
using System.Data;
using Zhuang.Data;

namespace Zhuang.NPOI.Test
{

    [TestClass]
    public class WorkbookBuilderTest
    {
        DbAccessor _dba = DbAccessor.Get();

        [TestMethod]
        public void HSSFWorkbookBuilderTest()
        {
            HSSFWorkbookBuilder builder = new HSSFWorkbookBuilder();
            DataTable dt = _dba.QueryDataTable("select top 10 * from sys_product");
            builder.SetDataTable(dt)
                .SetColumnCaption("产品名称", "ProductName").SetColumnWidth(20)
                .Build().Save(@"C:\npoitest.xls");
        }


        [TestMethod]
        public void XSSFWorkbookBuilderTest()
        {
            XSSFWorkbookBuilder builder = new XSSFWorkbookBuilder();
            DataTable dt = _dba.QueryDataTable("select * from sys_product");
            builder.SetDataTable(dt)
                .SetColumnCaption("产品名称", "ProductName").SetColumnWidth(20)
                .Build().Save(@"C:\npoitest.xlsx");
        }

        
    }
}
