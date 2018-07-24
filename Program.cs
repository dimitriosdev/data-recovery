using EFsetWidgetFix.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Windows.Forms;

namespace DataRecovery
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var serviceUrl = "https://api.efset.org/test-results/";
            var apiKey = "49a6f8e0-0384-4613-809e-bfa54a69dfd1";

            OpenFileDialog fd = new OpenFileDialog();
            fd.InitialDirectory = @"C:\Users\dimitrios.metozis\Downloads";
            fd.ShowDialog();

            if (fd.FileName.EndsWith("xlsx"))
            {
                var fi = new FileInfo(@fd.FileName);

                using (var p = new ExcelPackage(fi))
                {
                    var queryParamsList = new List<QueryParam>();
                    var ws = p.Workbook.Worksheets[1];
                    var firstRow = ExcelWorksheetExtension.GetRow(ws, 1);
                    var customerIdColumnIndex = GetColumnIndex(firstRow, "customer_id");
                    var insertDateColumnIndex = GetColumnIndex(firstRow, "insertdate_customer");
                    var vocabularyScoreColumnIndex = GetColumnIndex(firstRow, "Vocabulary Score");
                    var vocabularyPercentageColumnIndex = GetColumnIndex(firstRow, "Vocabulary Percentage");
                    var readingScoreColumnIndex = GetColumnIndex(firstRow, "Reading Score");
                    var readingPercentageColumnIndex = GetColumnIndex(firstRow, "Reading Percentage");
                    var listeningScoreColumnIndex = GetColumnIndex(firstRow, "Listening Score");
                    var listeningPercentageColumnIndex = GetColumnIndex(firstRow, "Listening Percentage");

                    var customerIdColumn = ExcelWorksheetExtension.GetColumn(ws, customerIdColumnIndex);
                    var customerIds = customerIdColumn.Skip(1);
                    foreach (var customerId in customerIds)
                    {
                        var queryParam = new QueryParam
                        {
                            CustomerId = customerId
                        };
                        queryParamsList.Add(queryParam);
                    }




                    var insertDateColumn = ExcelWorksheetExtension.GetColumn(ws, insertDateColumnIndex);
                    var insertDates = ExcelWorksheetExtension.ConvertToDate(insertDateColumn);
                    var insertDatesCollection = ExcelWorksheetExtension.GetDatesSet(insertDates);
                    var datesPairs = insertDatesCollection.AllKeys.SelectMany(insertDatesCollection.GetValues, (k, v) => new { key = k, value = v }).Skip(1);
                    for (int i = 0; i < datesPairs.Count(); i++)
                    {
                        queryParamsList[i].StartDate = datesPairs.ElementAt(i).key;
                        queryParamsList[i].EndDate = datesPairs.ElementAt(i).value;
                    }

                    int j = 2;
                    foreach (var queryParamsItem in queryParamsList)
                    {
                        string jsonString = QueryApi(apiKey, serviceUrl, queryParamsItem.CustomerId, queryParamsItem.StartDate, queryParamsItem.EndDate);
                        
                        if (HasScores(jsonString))
                        {
                            if (IsAdaptiveTest(jsonString))
                            {
                                var score_a = Newtonsoft.Json.JsonConvert.DeserializeObject<AdaptiveTestResult>(jsonString);
                                try
                                {
                                    ws.Cells[j, vocabularyPercentageColumnIndex+1].Value = score_a.AdaptiveResults.FirstOrDefault().Score.Combined;
                                    ws.Cells[j, vocabularyScoreColumnIndex+1].Value = score_a.AdaptiveResults.FirstOrDefault().Score.Cefr;
                                    ws.Cells[j, readingScoreColumnIndex+1].Value = score_a.AdaptiveResults.FirstOrDefault().Score.Reading.Cefr;
                                    ws.Cells[j, readingPercentageColumnIndex+1].Value = score_a.AdaptiveResults.FirstOrDefault().Score.Reading.Score;
                                    ws.Cells[j, listeningScoreColumnIndex+1].Value = score_a.AdaptiveResults.FirstOrDefault().Score.Listening.Cefr;
                                    ws.Cells[j, listeningPercentageColumnIndex+1].Value = score_a.AdaptiveResults.FirstOrDefault().Score.Listening.Score;

                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine(ex.Message);
                                }
                            }
                            else if (IsFixedTest(jsonString))
                            {
                                Console.Write("Fixed");
                            }
                        }
                        j++;
                    }


                    p.Save();
                }

            }
            Console.WriteLine("ww");
            Console.ReadLine();

        }


        private static int GetColumnIndex(string[] firstRow, string columnName)
        {
            for (int i = 0; i < firstRow.Length; i++)
            {
                if (string.Equals(firstRow[i].ToString(), columnName, StringComparison.InvariantCultureIgnoreCase))
                {
                    //Console.WriteLine(columnName + " column:" + i);
                    return i;
                }
            }
            //Console.WriteLine(columnName + " doesn't exist.");
            return -1;
        }

        private static string QueryApi(string aKey, string sUrl, string cId, string startDate, string endDate)
        {
            string scoreJson = "";
            using (WebClient c = new WebClient())
            {
                try
                {
                    c.Headers["X-API-KEY"] = aKey;
                    var address = string.Format(@"{0}{1}?from={2}&to={3}", sUrl, cId, startDate, endDate);
                    byte[] data = c.DownloadData(address);
                    System.Text.Encoding enc = System.Text.Encoding.ASCII;
                    scoreJson = enc.GetString(data);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(ex.Message);
                }
            }
            return scoreJson;

        }

        private static bool HasScores(string jsonstring)
        {
            return jsonstring.Contains("test_results");
        }

        private static bool IsAdaptiveTest(string jsonstring)
        {
            return jsonstring.Contains("efset1506");
        }

        private static bool IsFixedTest(string jsonstring)
        {
            return jsonstring.Contains("express1604");
        }
    }

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

    public class QueryParam
    {
        public string CustomerId { get; set; }

        public string StartDate { get; set; }

        public string EndDate { get; set; }
    }

}
