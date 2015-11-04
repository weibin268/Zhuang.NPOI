using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace Zhuang.NPOI.Excel
{

    public class BuildContext
    {
        public WorkbookBuilder WorkbookBuilder { get; set; }
        public IWorkbook Workbook { get; set; }
        public IRow Row { get; set; }
        public ICell Cell { get; set; }
    }

    public delegate void BuildEventHandler(BuildContext context);

    public abstract class WorkbookBuilder
    {
        #region variable
        DataTable _dataTable;
        Dictionary<string, string> _dicColumnCaption = new Dictionary<string, string>();
        Dictionary<string, int> _dicColumnWidth = new Dictionary<string, int>();
        Dictionary<string, int> _dicColumnOrdinal = new Dictionary<string, int>();
        IList<string> _lsRemoveColumn = new List<string>();
        IList<string> _lsShowColumn = new List<string>();
        int _defaultColumnWidth = -1;
        string _currentColumnName = string.Empty;
        bool _showHeadRow = true; 
        #endregion

        #region event
        public event BuildEventHandler OnHeadRowCreated;
        public WorkbookBuilder AddOnHeadRowCreated(BuildEventHandler onHeadRowCreated)
        {
            OnHeadRowCreated += onHeadRowCreated;
            return this;
        }
        public event BuildEventHandler OnHeadRowCellCreated;
        public WorkbookBuilder AddOnHeadRowCellCreated(BuildEventHandler onHeadRowCellCreated)
        {
            OnHeadRowCellCreated += onHeadRowCellCreated;
            return this;
        }
        public event BuildEventHandler OnRowCreated;
        public WorkbookBuilder AddOnRowCreated(BuildEventHandler onRowCreated)
        {
            OnRowCreated += onRowCreated;
            return this;
        }
        public event BuildEventHandler OnRowCellCreated;
        public WorkbookBuilder AddOnRowCellCreated(BuildEventHandler onRowCellCreated)
        {
            OnRowCellCreated += onRowCellCreated;
            return this;
        }
        #endregion

        public WorkbookBuilder()
        {

        }

        #region common setting
        public WorkbookBuilder SetDataTable(DataTable dt)
        {
            _dataTable = dt;
            return this;
        }

        public WorkbookBuilder RemoveColumn(string columnName)
        {
            _lsRemoveColumn.Add(columnName);
            return this;
        }

        public WorkbookBuilder SetShowColumns(IList<string> columnNames)
        {
            _lsShowColumn = columnNames;
            return this;
        }

        public WorkbookBuilder SetDefaultColumnWidth(int width)
        {
            _defaultColumnWidth = width;
            return this;
        }

        public WorkbookBuilder HideHeadRow()
        {
            _showHeadRow = false;
            return this;
        }
        #endregion

        #region column setting
        public WorkbookBuilder SetCurrentColumnName(string currentColumnName)
        {
            _currentColumnName = currentColumnName;
            return this;
        }

        public WorkbookBuilder SetColumnCaption(string columnCaption, string currentColumnName = null)
        {
            return SetColumn<string>(_dicColumnCaption, currentColumnName, columnCaption);
        }

        public WorkbookBuilder SetColumnWidth(int columnWidth, string currentColumnName = null)
        {
            return SetColumn<int>(_dicColumnWidth, currentColumnName, columnWidth);
        }

        public WorkbookBuilder SetColumnOrdinal(int columnOrdinal, string currentColumnName = null)
        {
            return SetColumn<int>(_dicColumnOrdinal, currentColumnName, columnOrdinal);
        }

        private WorkbookBuilder SetColumn<T>(Dictionary<string, T> dicSetting, string columnName, T settingValue)
        {

            if (columnName == null)
                columnName = _currentColumnName;

            if (columnName == string.Empty) return this;

            columnName = columnName.ToLower();

            _currentColumnName = columnName;


            if (dicSetting.ContainsKey(columnName))
            {
                dicSetting[columnName] = settingValue;
            }
            else
            {
                dicSetting.Add(columnName, settingValue);
            }

            return this;
        }
        #endregion
        
        protected abstract IWorkbook CreateWorkbook();

        public IWorkbook Build()
        {
            IWorkbook workbook = CreateWorkbook();
            var sheet = workbook.CreateSheet();
            int currentRowIndex = 0;

            if (_dataTable != null)
            {
                #region _lsShowColumn
                if (_lsShowColumn != null && _lsShowColumn.Count > 0)
                {
                    foreach (DataColumn dc in _dataTable.Columns)
                    {
                        if (!_lsShowColumn.Contains(dc.ColumnName))
                        {
                            _lsRemoveColumn.Add(dc.ColumnName);
                        }
                    }
                }
                #endregion

                #region _lsRemoveColumn
                foreach (string tempRemove in _lsRemoveColumn)
                {
                    if (_dataTable.Columns.IndexOf(tempRemove) >= 0)
                    {
                        _dataTable.Columns.Remove(tempRemove);
                    }
                }
                #endregion

                #region _dicColumnCaption
                foreach (DataColumn dc in _dataTable.Columns)
                {
                    string tempColumnName = dc.ColumnName.ToLower();

                    if (_dicColumnCaption.ContainsKey(tempColumnName))
                    {
                        dc.Caption = _dicColumnCaption[tempColumnName];
                    }
                }
                #endregion

                #region _dicColumnOrdinal
                foreach (var item in _dicColumnOrdinal)
                {
                    if (_dataTable.Columns.IndexOf(item.Key) >= 0)
                    {
                        _dataTable.Columns[item.Key].SetOrdinal(item.Value);
                    }
                }
                #endregion

                #region headRow
                if (_showHeadRow)
                {
                    var headRow = sheet.CreateRow(currentRowIndex++);

                    if (OnHeadRowCreated != null)
                        OnHeadRowCreated(new BuildContext() { WorkbookBuilder = this, Workbook = workbook, Row = headRow, Cell = null });

                    for (int i = 0; i < _dataTable.Columns.Count; i++)
                    {
                        var tempCoumnName = _dataTable.Columns[i].ColumnName;
                        ICell tempCell = headRow.CreateCell(i);
                        string tempCellvalue = _dataTable.Columns[i].Caption ?? tempCoumnName;
                        tempCell.SetCellValue(tempCellvalue);

                        SetColumnWidthByCell(tempCell, tempCoumnName);

                        if (OnHeadRowCellCreated != null)
                            OnHeadRowCellCreated(new BuildContext() { WorkbookBuilder = this, Workbook = workbook, Row = headRow, Cell = tempCell });
                    }
                }
                #endregion

                #region _dataTable.Rows
                foreach (DataRow dr in _dataTable.Rows)
                {
                    int currentColumnsIndex = 0;
                    var tempRow = sheet.CreateRow(currentRowIndex++);

                    if (OnRowCreated != null)
                        OnRowCreated(new BuildContext() { WorkbookBuilder = this, Workbook = workbook, Row = tempRow, Cell = null });

                    foreach (DataColumn dc in _dataTable.Columns)
                    {
                        var tempCell = tempRow.CreateCell(currentColumnsIndex++);
                        tempCell.SetCellValue(dr[dc.ColumnName].ToString());

                        if (!_showHeadRow && currentColumnsIndex == 1)
                        {
                            SetColumnWidthByCell(tempCell, dc.ColumnName);
                        }

                        if (OnRowCellCreated != null)
                            OnRowCellCreated(new BuildContext() { WorkbookBuilder = this, Workbook = workbook, Row = tempRow, Cell = tempCell });
                    }
                }
                #endregion

            }

            return workbook;
        }

        #region common methds
        private void SetColumnWidthByCell(ICell cell, string columnName)
        {
            if (_dicColumnWidth.ContainsKey(columnName.ToLower()))
            {
                cell.SetColumnWidth(_dicColumnWidth[columnName.ToLower()]);
            }
            else
            {
                if (_defaultColumnWidth > 0)
                    cell.SetColumnWidth(_defaultColumnWidth);
            }
        } 
        #endregion
    }
}
