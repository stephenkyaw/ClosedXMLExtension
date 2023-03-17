using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Web;

namespace ClosedXMLExtension.Extensions
{
    public static class ClosedXMLExtension
    {

        private const string DEFAULT_PASSWORD = "Admin@123!@#";

        public static void ProtectWithDefaultPassword(this IXLWorksheet worksheet)
        {
            worksheet.Protect(DEFAULT_PASSWORD);
        }

        /// <summary>
        /// Set DataValidation By TEntity FieldName
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="worksheet"></param>
        /// <param name="fieldName"></param>
        /// <param name="list"></param>
        /// <param name="inCellDropdown"></param>
        /// <returns></returns>
        public static IXLWorksheet SetDataValidation<TEntity>(this IXLWorksheet worksheet, string fieldName, string list, bool inCellDropdown) where TEntity : class, new()
        {
            List<string> columnNameList = new List<string>();
            Dictionary<int, string> requireColumnList = new Dictionary<int, string>();

            var properties = new TEntity().GetType().GetProperties();

            foreach (var property in properties)
            {
                if (!Attribute.IsDefined(property, typeof(ExceptClosedXMLColumnAttribute)))
                {
                    columnNameList.Add(property.Name);
                }
            }

            worksheet.Column(columnNameList.IndexOf(fieldName) + 1).SetDataValidation().List(list, inCellDropdown);

            return worksheet;
        }

        /// <summary>
        /// Set Data Validation By Other Sheets 
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="worksheet"></param>
        /// <param name="fieldName"></param>
        /// <param name="range"></param>
        /// <param name="inCellDropdown"></param>
        /// <returns></returns>
        public static IXLWorksheet SetDataValidation<TEntity>(this IXLWorksheet worksheet, string fieldName, IXLRange range, bool inCellDropdown) where TEntity : class, new()
        {
            List<string> columnNameList = new List<string>();

            var properties = new TEntity().GetType().GetProperties();

            foreach (var property in properties)
            {
                if (!Attribute.IsDefined(property, typeof(ExceptClosedXMLColumnAttribute)))
                {
                    columnNameList.Add(property.Name);
                }
            }

            worksheet.Column(columnNameList.IndexOf(fieldName) + 1)
                .SetDataValidation()
                .List(range, inCellDropdown);

            return worksheet;
        }


        /// <summary>
        /// TEntity Property Name Write To Excel Header
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="worksheets"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static IXLWorksheet Add<TEntity>(this IXLWorksheets worksheets, string sheetName) where TEntity : class, new()
        {
            List<string> columnNameList = new List<string>();
            List<string> requireColorColumns = new List<string>();

            var properties = new TEntity().GetType().GetProperties();

            var worksheet = worksheets.Add(sheetName);

            var style = XLWorkbook.DefaultStyle;
            style.Border.DiagonalUp = true;
            style.Border.DiagonalDown = true;
            style.Border.DiagonalBorder = XLBorderStyleValues.Thick;
            style.Border.DiagonalBorderColor = XLColor.Red;

            worksheet.Style = style;
            worksheet.SheetView.FreezeRows(1);

            foreach (var property in properties)
            {
                if (!Attribute.IsDefined(property, typeof(ExceptClosedXMLColumnAttribute)))
                {
                    columnNameList.Add(property.Name);

                }
                if (!Attribute.IsDefined(property, typeof(ExceptClosedXMLColumnAttribute)) &&
                    Attribute.IsDefined(property, typeof(RequireClosedXMLColumnAttribute)))
                {
                    requireColorColumns.Add(property.Name);
                }
            }

            foreach (var columnName in columnNameList)
            {
                int columnIndex = columnNameList.IndexOf(columnName) + 1;
                worksheet.Cell(1, columnIndex).Value = columnName.ToUpper();
                worksheet.Column(columnIndex).AdjustToContents();
            }

            foreach (var requireColor in requireColorColumns)
            {
                int columnIndex = columnNameList.IndexOf(requireColor) + 1;
                worksheet.Cell(1, columnIndex).Style.Font.FontColor = XLColor.Red;
                worksheet.Cell(1, columnIndex).Style.Font.Bold = true;
            }


            worksheet.Style.Font.SetFontSize(10);
            worksheet.Row(1).Style.Font.Bold = true;
            worksheet.Row(1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            worksheet.Row(1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
            worksheet.Row(1).Height = 25;

            return worksheet;
        }

        /// <summary>
        /// Set Cell Value by Column Name
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="columnName"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        public static IXLWorksheet SetCellValues(this IXLWorksheet worksheet, string columnName, List<string> values)
        {
            int count = 1;
            foreach (var value in values)
            {
                worksheet.Cell($"{columnName + count.ToString()}").SetValue<string>(value);
                worksheet.Cell($"{columnName + count.ToString()}").SetDataType(XLDataType.Text);
                count++;
            }

            worksheet.Style.Font.SetFontSize(10);
            worksheet.Column(columnName).AdjustToContents();

            return worksheet;
        }

        /// <summary>
        /// Get Cell Value List By FieldName
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="xLRangeRows"></param>
        /// <param name="fidelName"></param>
        /// <returns></returns>
        public static IEnumerable<string> GetCellValues<TEntity>(this IEnumerable<IXLRangeRow> xLRangeRows, string fidelName)
             where TEntity : class, new()
        {
            List<string> values = new List<string>();

            List<string> columnNameList = new List<string>();

            var properties = new TEntity().GetType().GetProperties();

            foreach (var property in properties)
            {
                if (!Attribute.IsDefined(property, typeof(ExceptClosedXMLColumnAttribute)))
                {
                    columnNameList.Add(property.Name);
                }
            }
            foreach (var row in xLRangeRows)
            {
                if (!row.IsEmpty())
                {
                    int columnIndex = columnNameList.IndexOf(fidelName) + 1;
                    values.Add(row.Cell(columnIndex).GetValue<string>());
                }
            }

            return values;
        }


        /// <summary>
        /// Get List<TEntity> From Excel
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="xLRangeRows"></param>
        /// <returns></returns>
        public static IEnumerable<TEntity> GetEntities<TEntity>(this IEnumerable<IXLRangeRow> xLRangeRows, out List<string> _errorMessages)
            where TEntity : class, new()
        {
            _errorMessages = new List<string>();

            var entities = new List<TEntity>();
            List<string> columnNameList = new List<string>();

            var properties = new TEntity().GetType().GetProperties();

            foreach (var property in properties)
            {
                if (!Attribute.IsDefined(property, typeof(ExceptClosedXMLColumnAttribute)))
                {
                    columnNameList.Add(property.Name);
                }
            }
            foreach (var row in xLRangeRows)
            {
                if (!row.IsEmpty())
                {
                    TEntity entity = new TEntity();

                    foreach (var column in columnNameList)
                    {
                        var property = properties.FirstOrDefault(x => x.Name.ToUpper() == column.ToUpper());
                        int columnIndex = columnNameList.IndexOf(column) + 1;
                        object cellValue = GetValue(row, columnIndex, property.PropertyType);

                        var attri = property.GetCustomAttribute<RequireClosedXMLColumnAttribute>();

                        if (attri != null && string.IsNullOrEmpty(Convert.ToString(cellValue)))
                        {
                            string errorMessage = $"Column name {property.Name} in row number {row.RowNumber()} is cannot be empty.";
                            _errorMessages.Add(errorMessage);
                        }
                        TrySetProperty(entity, column, cellValue);
                    }
                    entities.Add(entity);
                }
            }
            return entities;
        }


        private static object GetValue(IXLRangeRow row, int columnIndex, Type type)
        {
            if (type == typeof(string))
            {
                return row.Cell(columnIndex).GetValue<string>();
            }
            else if (type == typeof(int))
            {
                if (!row.Cell(columnIndex).IsNull() && !row.Cell(columnIndex).IsEmpty())
                {
                    return row.Cell(columnIndex).GetValue<int>();
                }
                else return null;
            }
            else if (type == typeof(DateTime))
            {
                if (!row.Cell(columnIndex).IsNull() && !row.Cell(columnIndex).IsEmpty())
                {
                    return row.Cell(columnIndex).GetValue<DateTime>();
                }
                else return null;
            }
            else if (type == typeof(Decimal))
            {
                if (!row.Cell(columnIndex).IsNull() && !row.Cell(columnIndex).IsEmpty())
                {
                    return row.Cell(columnIndex).GetValue<decimal>();
                }
                else return null;
            }
            else
            {
                return null;
            }
        }
        private static void TrySetProperty(object obj, string property, object value)
        {
            var prop = obj.GetType().GetProperty(property, BindingFlags.Public | BindingFlags.Instance);
            if (prop != null && prop.CanWrite)
            {
                prop.SetValue(obj, value, null);
            }
        }

    }
}