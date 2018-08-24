using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Globalization;
using System.Linq;

namespace DataRecoveryLib.Extensions
{
    public class ExcelExtension
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
            string[] formats = {"M/d/yyyy h:mm:ss tt", "M/d/yyyy h:mm tt",
                     "MM/dd/yyyy hh:mm:ss", "M/d/yyyy h:mm:ss",
                     "M/d/yyyy hh:mm tt", "M/d/yyyy hh tt",
                     "M/d/yyyy h:mm", "M/d/yyyy h:mm",
                     "MM/dd/yyyy hh:mm", "M/dd/yyyy hh:mm"};
            foreach (var rawDate in insertDateColumn)
            {
                DateTime formattedDate = DateTime.MinValue; ;
                DateTime.TryParseExact(rawDate, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out formattedDate);
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
