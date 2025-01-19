using System.Collections;
using ExcelReportGenerator.ExcelEntities;
using ExcelReportGenerator.Interfaces;

namespace ExcelReportGenerator
{
    public abstract class BaseExcelWriter
    {
        public abstract bool CellValueIsNumber(object value);
        public abstract string PrepareEnumValueToShow(object value);
        public abstract string PrepareDateTimeValueToShow(object value, IDateTimeProperties properties);
        public abstract string PrepareNumberValueToShow(object value, INumberProperties properties);

        public virtual List<ExcelRow> CreateExcelRows<T>(IEnumerable<T> dataList, List<IExcelColumn> columns)
            where T : class, IExcelData
        {
            var rows = new List<ExcelRow>();
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
                var row = new ExcelRow() { Id = id };
                foreach (var column in displayNormalColumns)
                {
                    var property = type.GetProperty(FirstCharToUpper(column.Name));
                    if (property == null)
                    {
                        row.Cells.Add(new ExcelCell() { Order = column.Order, Value = string.Empty });
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

                    row.Cells.Add(new ExcelCell() { Order = column.Order, Value = value });
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
                                row.Cells.Add(new ExcelCell() { Order = column.Order, Value = string.Empty });
                                continue;
                            }

                            var valueType = value.Value.GetType();

                            if (valueType == typeof(DateTime))
                            {
                                val = PrepareDateTimeValueToShow(value, column.DataProperties.DateTimeProperties);
                            }

                            row.Cells.Add(new ExcelCell() { Order = column.Order, Value = val });
                        }
                    }
                }


                if (horizontal != null)
                {
                    var value = type.GetProperty(FirstCharToUpper(horizontal.Name))?.GetValue(item);
                    row.HorizontalCells.Add(new ExcelCell() { Order = horizontal.Order, Value = value });
                }

                rows.Add(row);
            }

            return rows;

        }

        private static string FirstCharToUpper(string input)
        {
            if (string.IsNullOrEmpty(input))
                return string.Empty;
            return input.First().ToString().ToUpper() + string.Join("", input.Skip(1));
        }

    }
}
