using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace Zhuang.NPOI.Excel
{

    public class AdaptContext
    {
        public WorkbookDataAdapter WorkbookDataAdapter { get; set; }
        public IWorkbook Workbook { get; set; }
        public IRow Row { get; set; }
        public ICell Cell { get; set; }
        public object DataRow { get; set; }
        public object DataCell { get; set; }
    }

    public delegate bool AdaptEventHandler(AdaptContext context);

    public class WorkbookDataAdapter
    {
        #region variable
        IWorkbook _workbook;
        bool _includeHeadRow = false;
        Dictionary<string, string> _dicColumnNameMapping = new Dictionary<string, string>();
        #endregion

        #region event
        public event AdaptEventHandler OnRowCellAdapt;
        public WorkbookDataAdapter AddOnRowCellAdapt(AdaptEventHandler onRowCellAdapt)
        {
            OnRowCellAdapt += onRowCellAdapt;
            return this;
        }

        public event AdaptEventHandler OnRowAdapt;
        public WorkbookDataAdapter AddOnRowAdapt(AdaptEventHandler onRowAdapt)
        {
            OnRowAdapt += onRowAdapt;
            return this;
        }
        #endregion

        public WorkbookDataAdapter(IWorkbook workbook)
        {
            _workbook = workbook;
        }

        #region setting
        public WorkbookDataAdapter SetIncludeHeadRow(bool includeHeadRow)
        {
            _includeHeadRow = includeHeadRow;
            return this;
        }

        public WorkbookDataAdapter AddColumnNameMapping(string excelColumnName, string dataColumnName)
        {
            if (_dicColumnNameMapping.ContainsKey(excelColumnName))
            {
                _dicColumnNameMapping[excelColumnName] = dataColumnName;
            }
            else
            {
                _dicColumnNameMapping.Add(excelColumnName, dataColumnName);
            }
            return this;
        }

        public WorkbookDataAdapter SetColumnNameMapping(Dictionary<string, string> dicColumnNameMapping)
        {
            _dicColumnNameMapping = dicColumnNameMapping;
            return this;
        }
        #endregion

        #region ToData
        public DataTable ToDataTable()
        {
            DataTable dtResult = new DataTable();
            ISheet sheet = _workbook.GetSheetAt(0);
            IRow headerRow = sheet.GetRow(0);
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();

            int colCount = headerRow.LastCellNum;
            int rowCount = sheet.LastRowNum;

            for (int i = 0; i < colCount; i++)
            {
                dtResult.Columns.Add(GetDataColumnName(headerRow.GetCell(i).ToString()));
            }

            if (!_includeHeadRow)
            {
                rows.MoveNext();
            }

            while (rows.MoveNext())
            {
                IRow row = (IRow)rows.Current;
                DataRow dr = dtResult.NewRow();

                for (int i = 0; i < colCount; i++)
                {
                    ICell cell = row.GetCell(i);

                    if (cell != null)
                    {
                        dr[i] = cell.ToString();
                        if (OnRowCellAdapt != null)
                        {
                            if (!OnRowCellAdapt(new AdaptContext()
                            {
                                WorkbookDataAdapter = this,
                                Workbook = _workbook,
                                Row = row,
                                Cell = cell,
                                DataCell = dr[i]
                            }))
                            {
                                goto RowEnd;
                            }
                        }
                    }
                }

                if (OnRowAdapt != null)
                {
                    if (!OnRowAdapt(new AdaptContext()
                    {
                        WorkbookDataAdapter = this,
                        Workbook = _workbook,
                        Row = row,
                        DataRow = dr
                    }))
                    {
                        goto RowEnd;
                    }
                }

                dtResult.Rows.Add(dr);

                RowEnd:;
            }

            return dtResult;
        }

        public IList<T> ToList<T>()
        {
            IList<T> lsResult = new List<T>();

            ISheet sheet = _workbook.GetSheetAt(0);
            IRow headerRow = sheet.GetRow(0);
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();

            int colCount = headerRow.LastCellNum;
            int rowCount = sheet.LastRowNum;

            if (!_includeHeadRow)
            {
                rows.MoveNext();
            }

            while (rows.MoveNext())
            {
                IRow row = (IRow)rows.Current;
                T entity = (T)Activator.CreateInstance(typeof(T));

                for (int i = 0; i < colCount; i++)
                {
                    ICell cell = row.GetCell(i);

                    if (cell != null)
                    {
                        string colName = GetDataColumnName(headerRow.GetCell(i).ToString());
                        var pi = entity.GetType().GetProperty(colName);
                        if (pi != null)
                        {
                            pi.SetValue(entity, Convert.ChangeType(cell.ToString(), pi.PropertyType), null);

                            if (OnRowCellAdapt != null)
                            {
                                if (!OnRowCellAdapt(new AdaptContext()
                                {
                                    WorkbookDataAdapter = this,
                                    Workbook = _workbook,
                                    Row = row,
                                    Cell = cell,
                                    DataCell = pi.GetValue(entity, null)
                                }))
                                {
                                    goto RowEnd;
                                }
                            }
                        }
                    }
                    //dr[i] = cell.ToString();
                }

                if (OnRowAdapt != null)
                {
                    if (!OnRowAdapt(new AdaptContext()
                    {
                        WorkbookDataAdapter = this,
                        Workbook = _workbook,
                        Row = row,
                        DataRow = entity
                    }))
                    {
                        goto RowEnd;
                    }
                }

                lsResult.Add(entity);

                RowEnd:;
            }

            return lsResult;
        }
        #endregion

        #region common methds
        private string GetDataColumnName(string excelColumnName)
        {
            excelColumnName = excelColumnName.Trim();

            if (_dicColumnNameMapping.ContainsKey(excelColumnName))
            {
                return _dicColumnNameMapping[excelColumnName];
            }
            else
            {
                return excelColumnName;
            }

        } 
        #endregion
    }
}
