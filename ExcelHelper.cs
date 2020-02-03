using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace Helpers
{
    public static class ExcelHelper
    {
        private static string GetValueAsString(this ICell cell)
        {
            cell.SetCellType(CellType.String);
            return cell.StringCellValue;
        }

        private static int GetValueAsInteger(this ICell cell)
        {
            cell.SetCellType(CellType.Numeric);
            return cell.NumericCellValue.ToInteger();
        }

        private static decimal GetValueAsDecimal(this ICell cell)
        {
            cell.SetCellType(CellType.Numeric);
            return cell.NumericCellValue.ToDecimal();
        }

        private static bool GetValueAsBoolean(this ICell cell)
        {
            cell.SetCellType(CellType.Boolean);
            return cell.BooleanCellValue;
        }

        private static DateTime GetValueAsDateTime(this ICell cell)
        {
            return cell.DateCellValue;
        }

        private static DateTime? GetValueAsNullableDateTime(this ICell cell)
        {
            try
            {
                var date = cell.DateCellValue;
                if (date != default)
                    return date;
                else return null;
            }
            catch
            {
                return null;
            }
        }

        public static List<T> Import<T>(string filePath) where T : new()
        {
            XSSFWorkbook xssfwb;
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                xssfwb = new XSSFWorkbook(stream);
            }

            var sheet = xssfwb.GetSheetAt(0);

            var rowHeader = sheet.GetRow(0);
            var colIndexList = new Dictionary<string, int>();
            foreach (var cell in rowHeader.Cells)
            {
                var colName = GetColumnName(cell.StringCellValue);
                colIndexList.Add(colName, cell.ColumnIndex);
            }

            var listResult = new List<T>();
            var currentRow = 1;
            while (currentRow <= sheet.LastRowNum)
            {
                var row = sheet.GetRow(currentRow);
                var obj = new T();

                foreach (var property in typeof(T).GetProperties())
                {
                    var colName = GetColumnName(property.Name);
                    if (!colIndexList.ContainsKey(colName)) continue;

                    var colIndex = colIndexList[colName];
                    var cell = row.GetCell(colIndex);

                    if (cell == null)
                        property.SetValue(obj, null);
                    else if (property.PropertyType == typeof(string))
                        property.SetValue(obj, cell.GetValueAsString());
                    else if (property.PropertyType == typeof(int))
                        property.SetValue(obj, cell.GetValueAsInteger());
                    else if (property.PropertyType == typeof(int?))
                        property.SetValue(obj, cell.GetValueAsInteger());
                    else if (property.PropertyType == typeof(decimal))
                        property.SetValue(obj, cell.GetValueAsDecimal());
                    else if (property.PropertyType == typeof(decimal?))
                        property.SetValue(obj, cell.GetValueAsDecimal());
                    else if (property.PropertyType == typeof(DateTime))
                        property.SetValue(obj, cell.GetValueAsDateTime());
                    else if (property.PropertyType == typeof(DateTime?))
                        property.SetValue(obj, cell.GetValueAsNullableDateTime());
                    else if (property.PropertyType == typeof(bool))
                        property.SetValue(obj, cell.GetValueAsBoolean());
                    else if (property.PropertyType == typeof(bool?))
                        property.SetValue(obj, cell.GetValueAsBoolean());
                    else if (property.PropertyType == typeof(Guid))
                        property.SetValue(obj, cell.GetValueAsString().ToGuid().GetValueOrDefault());
                    else if (property.PropertyType == typeof(Guid?))
                        property.SetValue(obj, cell.GetValueAsString().ToGuid());
                    else
                        property.SetValue(obj, Convert.ChangeType(cell.GetValueAsString(), property.PropertyType));
                }

                listResult.Add(obj);
                currentRow++;
            }

            return listResult;
        }

        private static string GetColumnName(string value)
        {
            return value.Replace(" ", "").ToLower();
        }

        public static byte[] CreateTemplateFile<T>()
        {
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("Sheet1");
            var row = sheet.CreateRow(0);

            var properties = typeof(T).GetProperties();

            var font = workbook.CreateFont();
            font.IsBold = true;
            var style = workbook.CreateCellStyle();
            style.SetFont(font);

            var colIndex = 0;
            foreach (var property in properties)
            {
                var cell = row.CreateCell(colIndex);
                cell.SetCellValue(property.Name);
                cell.CellStyle = style;
                colIndex++;
            }

            var stream = new MemoryStream();
            workbook.Write(stream);
            var content = stream.ToArray();

            return content;
        }

        public static byte[] CreateFile<T>(List<T> source)
        {
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("Sheet1");
            var rowHeader = sheet.CreateRow(0);

            var properties = typeof(T).GetProperties();

            //header
            var font = workbook.CreateFont();
            font.IsBold = true;
            var style = workbook.CreateCellStyle();
            style.SetFont(font);

            var colIndex = 0;
            foreach (var property in properties)
            {
                var cell = rowHeader.CreateCell(colIndex);
                cell.SetCellValue(property.Name);
                cell.CellStyle = style;
                colIndex++;
            }
            //end header


            //content
            var rowNum = 1;
            foreach (var item in source)
            {
                var rowContent = sheet.CreateRow(rowNum);

                var colContentIndex = 0;
                foreach (var property in properties)
                {
                    var cellContent = rowContent.CreateCell(colContentIndex);
                    var value = property.GetValue(item, null);

                    if (value == null)
                    {
                        cellContent.SetCellValue("");
                    }
                    else if (property.PropertyType == typeof(string))
                    {
                        cellContent.SetCellValue(value.ToString());
                    }
                    else if (property.PropertyType == typeof(int) || property.PropertyType == typeof(int?))
                    {
                        cellContent.SetCellValue(value.ToInteger());
                    }
                    else if (property.PropertyType == typeof(decimal) || property.PropertyType == typeof(decimal?))
                    {
                        cellContent.SetCellValue((double)value.ToDecimal());
                    } 
                    else if (property.PropertyType == typeof(DateTime) || property.PropertyType == typeof(DateTime?))
                    {
                        var dateValue = (DateTime)value;
                        cellContent.SetCellValue(dateValue.ToString("yyyy-MM-dd"));
                    }
                    else cellContent.SetCellValue(value.ToString());

                    colContentIndex++;
                }

                rowNum++;
            }

            //end content


            var stream = new MemoryStream();
            workbook.Write(stream);
            var content = stream.ToArray();

            return content;
        }
    }
}
