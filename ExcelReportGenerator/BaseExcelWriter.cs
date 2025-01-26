using System.Collections;
using ExcelReportGenerator.ExcelEntities;
using ExcelReportGenerator.Interfaces;
using OfficeOpenXml;

namespace ExcelReportGenerator
{
    public abstract class BaseExcelWriter
    {
        protected BaseExcelWriter()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public abstract bool CellValueIsNumber(object value);
        public abstract string PrepareEnumValueToShow(object value);
        public abstract string PrepareDateTimeValueToShow(object value, IDateTimeProperties properties);
        public abstract string PrepareNumberValueToShow(object value, INumberProperties properties);

        public virtual List<ExcelReportRow> CreateExcelRows<T>(IEnumerable<T> dataList, List<IExcelReportColumn> columns)
            where T : class, IExcelData
        {
            var rows = new List<ExcelReportRow>();
            var type = typeof(T);

            var displayNormalColumns = columns
                .Where(a => !a.Excluded && !a.IsHorizontal && !a.IsDynamicColumn)
                .OrderBy(a => a.Order).ToList();

            var displayDynamicColumns = columns
                .Where(a => !a.Excluded && !a.IsHorizontal && a.IsDynamicColumn)
                .OrderBy(a => a.Order).ToList();

            var horizontal = columns.FirstOrDefault(a => !a.Excluded && a.IsHorizontal);

            foreach (var item in dataList)
            {
                var id = type.GetProperty("Id")?.GetValue(item);
                var row = new ExcelReportRow() { Id = id };
                foreach (var column in displayNormalColumns)
                {
                    var property = type.GetProperty(column.Name.FirstCharToUpper());
                    if (property == null)
                    {
                        row.Cells.Add(new ExcelReportCell() { Order = column.Order, Value = string.Empty });
                        continue;
                    }

                    var value = property.GetValue(item);
                    var propertyType = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;

                    if (propertyType.IsEnum && value != null)
                    {
                        value = PrepareEnumValueToShow(value);
                    }

                    if (propertyType == typeof(DateTime) && value!=null)
                    {
                        value = PrepareDateTimeValueToShow(value, column.DataProperties.DateTimeProperties);
                    }

                    if (propertyType == typeof(string) && value != null && CellValueIsNumber(value))
                    {
                        value = PrepareNumberValueToShow(value, column.DataProperties.NumberProperties);
                    }

                    row.Cells.Add(new ExcelReportCell() { Order = column.Order, Value = value });
                }

                if (displayDynamicColumns.Any())
                {
                    var properties = type.GetProperties();

                    var listProperty = properties.FirstOrDefault(property =>
                         property.PropertyType.IsGenericType && property.PropertyType.GetGenericTypeDefinition() ==
                         typeof(List<>));
                    if (listProperty != null)
                    {
                        var dynamicValues = (IList)listProperty.GetValue(item)!;

                        foreach (var dynamicValue in dynamicValues)
                        {
                            var value = (IExcelDynamicData)dynamicValue;
                            var column = displayDynamicColumns.FirstOrDefault(a => a.Name == value.Name);
                            if (column == null) continue;

                            var val = value.Value;
                            if (val == null)
                            {
                                row.Cells.Add(new ExcelReportCell() { Order = column.Order, Value = string.Empty });
                                continue;
                            }

                            var valueType = value.Value.GetType();

                            if (valueType == typeof(DateTime))
                            {
                                val = PrepareDateTimeValueToShow(value, column.DataProperties.DateTimeProperties);
                            }

                            row.Cells.Add(new ExcelReportCell() { Order = column.Order, Value = val });
                        }
                    }
                }


                if (horizontal != null)
                {
                    var value = type.GetProperty(horizontal.Name.FirstCharToUpper())?.GetValue(item);
                    row.HorizontalCells.Add(new ExcelReportCell() { Order = horizontal.Order, Value = value });
                }

                rows.Add(row);
            }

            return rows;

        }

        public ExcelReportRow CreateTotalSumRow<T>( T model, List<IExcelReportColumn> columns)
            where T : class, IExcelTotalDataModel
        {
            var row = new ExcelReportRow();
            var type = typeof(T);

            var columnsWithTotal = columns.Where(a => !a.Excluded && !a.IsHorizontal && a.HasTotalSum)
                .OrderBy(a => a.Order)
                .ToList();

            foreach (var column in columnsWithTotal)
            {
                var desiredName = $"Total{column.Name.FirstCharToUpper()}";
                var property = type.GetProperty(desiredName);
                if (property == null)
                {
                    row.Cells.Add(new ExcelReportCell() { Order = column.Order, Value = string.Empty });
                    continue;
                }

                var value = property.GetValue(model);
                var propertyType = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;


                if (propertyType.IsEnum && value != null)
                {
                    value = PrepareEnumValueToShow(value);
                }

                row.Cells.Add(new ExcelReportCell() { Order = column.Order, Value = value });
            }

            return row;
        }

        public byte[] GenerateExcelReport<T>(ISingleSheetExcelRequest<T> request) 
            where T : class, IExcelData
        {
            var excelRows = CreateExcelRows(request.DataList, request.ExcelProfile.Columns);

            using var package = new ExcelPackage();
            var sheet = package
                .CreateSheet(request.SheetName)
                .SetSheetDefaults(request.ExcelProfile.DefaultProperties); 

            return package.GetAsByteArray();
        }

        public byte[] GenerateExcelReport<T,TN>(ISingleSheetExcelRequestWithTotalRow<T,TN> request) 
            where T : class, IExcelData 
            where TN : class, IExcelTotalDataModel
        {
            var excelRows = CreateExcelRows(request.DataList, request.ExcelProfile.Columns);
            var totalSumRow = CreateTotalSumRow(request.TotalDataModel, request.ExcelProfile.Columns);

            using var package = new ExcelPackage();
            var sheet = package
                .CreateSheet(request.SheetName)
                .SetSheetDefaults(request.ExcelProfile.DefaultProperties)
                .AddHeaderRow(request.ExcelProfile);
            
            
            return package.GetAsByteArray();
        }
    }
}
