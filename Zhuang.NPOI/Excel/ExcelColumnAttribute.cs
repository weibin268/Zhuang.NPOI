using System;
using System.Collections.Generic;
using System.Text;

namespace Zhuang.NPOI.Excel
{

    [System.AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = true)]
    public sealed class ExcelColumnAttribute : Attribute
    {
        readonly string _name;

        public ExcelColumnAttribute(string name)
        {
            _name = name;

        }
        public string Name
        {
            get { return _name; }
        }

        public static Dictionary<string, string> GetColumnNameMapping(Type entityType)
        {
            Dictionary<string, string> dicResult = new Dictionary<string, string>();

            foreach (var pi in entityType.GetProperties())
            {
                var attributes = pi.GetCustomAttributes(typeof(ExcelColumnAttribute), false);
                if (attributes.Length < 1) continue;

                dicResult.Add(((ExcelColumnAttribute)(attributes[0])).Name, pi.Name);
            }

            return dicResult;
        }
    }
}
