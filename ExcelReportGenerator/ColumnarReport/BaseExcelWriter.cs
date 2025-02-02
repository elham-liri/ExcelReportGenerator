using OfficeOpenXml;
using System.Text.RegularExpressions;

namespace ExcelReportGenerator.ColumnarReport
{
    public abstract class BaseExcelWriter
    {
        protected BaseExcelWriter()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public virtual bool CellValueIsNumber(object value)
        {
            return value != null 
                   && !string.IsNullOrEmpty(value.ToString()) 
                   && Regex.IsMatch(value.ToString()!, @"^-?\d+(?:\.\d+)?$");
        }

        public abstract string PrepareNumberValueToShow(object value);
        public abstract string PrepareEnumValueToShow(object value);
        public abstract string PrepareDateTimeValueToShow(object value);


        public virtual IEnumerable<IEnumerable<(int order, object? value)>> PrepareExcelRows<T>(IEnumerable<T> dataSource,
            IEnumerable<IExcelReportColumn> columns) where T : class, IExcelData
        {
            var rows = new List<IEnumerable<(int order, object? value)>>();
            var type = typeof(T);

            var normalColumns = columns.Where(a => !a.NeedsCalculation).OrderBy(a => a.Order).ToList();

            foreach (var dataRow in dataSource)
            {
                var row = new List<(int order, object? value)>();
                foreach (var column in normalColumns)
                {
                    var property = type.GetProperty(FirstCharToUpper(column.SourceName));
                    if (property == null)
                    {
                        row.Add((column.Order, string.Empty));
                        continue;
                    }

                    var value = property.GetValue(dataRow);
                    var propertyType = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;

                    if (propertyType.IsEnum && value != null)
                    {
                        value = PrepareEnumValueToShow(value);
                    }
                    else if (propertyType == typeof(DateTime) && value != null)
                    {
                        value = PrepareDateTimeValueToShow(value);
                    }
                    else if (propertyType == typeof(string) && value != null && CellValueIsNumber(value))
                    {
                        value = PrepareNumberValueToShow(value);
                    }

                    row.Add((column.Order,value));
                }
                rows.Add(row);
            }
            return rows;
        }

        public virtual byte[] GenerateExcelFileBytes(IEnumerable<IEnumerable<(int order, object? value)>> excelRows,
            IExcelReportProfile excelProfile, string sheetName)
        {
            using var package = new ExcelPackage();
            var sheet = package
                .CreateSheet(sheetName)
                .SetSheetDefaultProperties(excelProfile.DefaultProperties);

            return package.GetAsByteArray();

        }

        internal string FirstCharToUpper(string input)
        {
            if (string.IsNullOrEmpty(input))
                return string.Empty;
            return input.First().ToString().ToUpper() + string.Join("", input.Skip(1));
        }
    }
}
