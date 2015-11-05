using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Zhuang.NPOI.Excel;
using NPOI.HSSF.UserModel;
using System.IO;
using System.Data;
using Zhuang.Data.Utility;
using System.Collections.Generic;
using Zhuang.NPOI.Models;

namespace Zhuang.NPOI.Test
{
    /// <summary>
    /// Excel导入测试
    /// </summary>
    [TestClass]
    public class WorkbookDataAdapterTest
    {
        [TestMethod]
        public void ToDataTableTest()
        {
            DataTable dt = new WorkbookDataAdapter(new HSSFWorkbook(new FileStream(@"C:\npoitest.xls",FileMode.Open,FileAccess.Read)))
               .AddOnRowAdapt(c =>
               {
                   bool result = true;

                   return result;
               }).ToDataTable();

            Console.WriteLine(DataTableUtil.ToString(dt));
        }

        [TestMethod]
        public void ToListTest()
        {
            IList<SysProduct> products = new WorkbookDataAdapter(new HSSFWorkbook(new FileStream(@"C:\npoitest.xls", FileMode.Open, FileAccess.Read)))
                .SetColumnNameMapping(ExcelColumnAttribute.GetColumnNameMapping(typeof(SysProduct)))
                .AddOnRowAdapt(c =>
                {
                    bool result = true;
                    var record = (SysProduct)c.DataRow;

                    return result;
                }).ToList<SysProduct>();

            foreach (var item in products)
            {
                Console.WriteLine(item.ProductCode+"|"+item.ProductName);
            }
        }
    }
}
