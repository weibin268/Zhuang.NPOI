using System;
using System.Collections.Generic;
using System.Text;

namespace Zhuang.NPOI.Utility
{
    public class ReflectionHelper
    {
        //public static string GetPropertyName<T>(Expression<Func<T, object>> expr)
        //{
        //    var rtn = "";
        //    if (expr.Body is UnaryExpression)
        //    {
        //        rtn = ((MemberExpression)((UnaryExpression)expr.Body).Operand).Member.Name;
        //    }
        //    else if (expr.Body is MemberExpression)
        //    {
        //        rtn = ((MemberExpression)expr.Body).Member.Name;
        //    }
        //    else if (expr.Body is ParameterExpression)
        //    {
        //        rtn = ((ParameterExpression)expr.Body).Type.Name;
        //    }
        //    return rtn;
        //}


        //public static string GetExcelColumnName<T>(Expression<Func<T, object>> expr)
        //{
        //    return ExcelColumnAttribute.GetColumnNameMapping(typeof(T)).First(m => m.Value == GetPropertyName<T>(expr)).Key;
        //}
    }
}
