using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Zhuang.Data;
using Zhuang.NPOI.Excel;

namespace Zhuang.NPOI.Test
{
    [TestClass]
    public class ImportToDbTest
    {

        [TestMethod]
        public void Test()
        {

            string filePath = @"C:\Users\zwb\Desktop\cartype.xlsx";

            using (var fs=new FileStream(filePath,FileMode.Open))
            { 

                var dt = new WorkbookDataAdapter(new XSSFWorkbook(fs))
                //.SetColumnNameMapping(dic)
                .AddOnRowAdapt(c =>
                {
                    return true;

                }).ToDataTable();

                DbAccessor dba = DbAccessor.Get();

                foreach (DataRow dr in dt.Rows)
                {
                    foreach (DataColumn dc in dt.Columns)
                    {
                        //if (dc.ColumnName.ToUpper() == "STATUS" || dc.ColumnName.ToUpper()== "BELONGTTYPE")
                        //{
                        //    dr[dc.ColumnName] = Int16.Parse(dr[dc.ColumnName].ToString());
                        //}


                        if (string.IsNullOrEmpty(dr[dc.ColumnName].ToString()))
                        {
                            dr[dc.ColumnName] = DBNull.Value;
                        }


                    }
                }

                dt.TableName = "Biz_Cartype";
                dba.BulkWriteToServer(dt);


                Console.WriteLine();

            }



        }
    }
}
