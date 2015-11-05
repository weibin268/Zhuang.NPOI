using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Zhuang.NPOI.Excel;
using NPOI.HSSF.UserModel;
using System.IO;
using System.Data;
using Zhuang.Data.Utility;

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

        }
    }
}
