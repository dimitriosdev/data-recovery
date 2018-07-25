using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Globalization;
using System.Linq;

namespace DataRecovery.Extensions
{
    public class ExcelWorksheetExtension
    {
        internal static string[] GetRow(ExcelWorksheet sheet, int rowNumber)
        {
            return sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, rowNumber + 1, sheet.Dimension.End.Column]
                .Select(cell => cell.GetValue<string>()).ToArray();
        }

        internal static string[] GetColumn(ExcelWorksheet sheet, int columnNumber)
        {
            List<string> columnValues = new List<string>();
            foreach (var cell in sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column + columnNumber, sheet.Dimension.End.Row, columnNumber + 1])
                columnValues.Add(cell.GetValue<string>());
            return columnValues.ToArray();
        }

        internal static DateTime[] ConvertToDate(string[] insertDateColumn)
        {
            var collection = new List<DateTime>();

            foreach (var rawDate in insertDateColumn)
            {
                DateTime formattedDate;
                DateTime.TryParseExact(rawDate.Substring(0, rawDate.Length - 3), "M/d/yyyy H:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out formattedDate);
                collection.Add(formattedDate);
            }

            return collection.ToArray();
        }

        internal static NameValueCollection GetDatesSet(DateTime[] insertDates)
        {
            NameValueCollection collection = new NameValueCollection();

            foreach (var insertDate in insertDates)
            {
                if (insertDate.Year > 1900)
                    collection.Add(insertDate.AddDays(-10).ToString("yyyy-MM-ddThh:mm:ss.000Z"), insertDate.AddDays(10).ToString("yyyy-MM-ddThh:mm:ss.000Z"));
                else
                    collection.Add(insertDate.ToString("yyyy-MM-ddThh:mm:ss.000Z"), insertDate.ToString("yyyy-MM-ddThh:mm:ss.000Z"));
            }
            return collection;
        }
    }
}
