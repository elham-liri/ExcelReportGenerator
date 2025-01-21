﻿using ExcelReportGenerator.Interfaces;

namespace ExcelReportGenerator.ExcelEntities
{
    public abstract class BaseExcelProfile:IBaseExcelProfile
    {
        public List<IExcelColumn> Columns { get; set; } = null!;
        public IExcelProfileDefaultProperties? DefaultProperties { get; set; }
        public IExcelProfileCellStyle? CellStyle { get; set; }
        public IExcelProfileHeaderStyle? HeaderStyle { get; set; }
        public IExcelProfileTotalSumStyle? TotalSumStyle { get; set; }

        public abstract void InitializeColumns();

        public virtual void ResetColumns(IEnumerable<IExcelColumn> columns)
        {
            Columns = columns.ToList();
        }

        public virtual void SetDefaultProperties()
        {
            DefaultProperties=new ExcelProfileDefaultProperties();
        }

        public virtual void SetHeaderStyle()
        {
            HeaderStyle=new ExcelProfileHeaderStyle();
        }

        public virtual void SetTotalSumStyle()
        {
            TotalSumStyle=new ExcelProfileTotalSumStyle();
        }

        public virtual void SetCellStyle()
        {
            CellStyle=new ExcelProfileCellStyle();
        }

        public virtual void AddDynamicColumns(IEnumerable<IExcelColumn> dynamicColumns)
        {
            dynamicColumns = dynamicColumns.OrderBy(a => a.Order);
            var existedColumnsCount = Columns.Count;
            foreach (var dynamicColumn in dynamicColumns)
            {
                dynamicColumn.Order = ++existedColumnsCount;
                dynamicColumn.IsDynamicColumn = true;
                Columns.Add(dynamicColumn);
            }
        }

        public List<IExcelColumn> GetDisplayedVerticalColumns()
        {
            return Columns.Where(a => !a.Excluded && !a.IsHorizontal).OrderBy(a => a.Order).ToList();
        }

        public List<IExcelColumn> GetDisplayedHorizontalColumns()
        {
            return Columns.Where(a => !a.Excluded && a.IsHorizontal).OrderBy(a => a.Order).ToList(); 
        }
    }
}
