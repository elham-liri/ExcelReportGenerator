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


        public virtual IEnumerable<TR> PrepareExcelRows<T, TR, TM>(IEnumerable<T> dataSource,
            IEnumerable<IExcelReportColumn> columns)
            where T : class, IExcelData
            where TR : class, IReportRow<TM>, new()
            where TM : class, IReportCell, new()
        {
            var rows = new List<TR>();
            var type = typeof(T);

            var normalColumns = columns.Where(a => !a.NeedsCalculation).OrderBy(a => a.Order).ToList();

            foreach (var dataRow in dataSource)
            {
                var row = new TR() { Cells = new List<TM>() };
                foreach (var column in normalColumns)
                {
                    var property = type.GetProperty(FirstCharToUpper(column.SourceName));
                    if (property == null)
                    {
                        row.Cells.Add(new TM() { Order = column.Order, Value = string.Empty });
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

                    row.Cells.Add(new TM() { Order = column.Order, Value = value });
                }

                rows.Add(row);
            }

            return rows;
        }

        public virtual byte[] GenerateExcelFileBytes<T, TC>(IEnumerable<T> excelRows,
            IExcelReportProfile excelProfile, string sheetName)
            where T : class, IReportRow<TC>, new()
            where TC : class, IReportCell, new()
        {
            using var package = new ExcelPackage();
            var sheet = package
                .CreateSheet(sheetName)
                .SetSheetDefaultProperties(excelProfile.DefaultProperties)
                .AddHeaderRows(excelProfile.Headers,excelProfile.Columns);

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
