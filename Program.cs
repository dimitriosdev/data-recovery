using DataRecovery.Extensions;
using DataRecovery.Models;
using DataRecoveryLib;
using EFsetWidgetFix.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;

namespace DataRecovery
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var validator = new FileValidator();
            

            var serviceUrl = "https://api.efset.org/test-results/";
            var apiKey = "49a6f8e0-0384-4613-809e-bfa54a69dfd1";

            

            if (validator.IsValidExcelFile())
            {


                var fi = new FileInfo(validator.FileName);

                using (var p = new ExcelPackage(fi))
                {
                    var queryParamsList = new List<RequestParameters>();
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
                        var queryParam = new RequestParameters
                        {
                            UserId = customerId
                        };
                        queryParamsList.Add(queryParam);
                    }




                    var insertDateColumn = ExcelWorksheetExtension.GetColumn(ws, insertDateColumnIndex);
                    var insertDates = ExcelWorksheetExtension.ConvertToDate(insertDateColumn);
                    var insertDatesCollection = ExcelWorksheetExtension.GetDatesSet(insertDates);
                    var datesPairs = insertDatesCollection.AllKeys.SelectMany(insertDatesCollection.GetValues, (k, v) => new { key = k, value = v }).Skip(1);
                    for (int i = 0; i < datesPairs.Count(); i++)
                    {
                        queryParamsList[i].From = datesPairs.ElementAt(i).key;
                        queryParamsList[i].To = datesPairs.ElementAt(i).value;
                    }

                    int j = 2;
                    foreach (var queryParamsItem in queryParamsList)
                    {
                        string jsonString = QueryApi(apiKey, serviceUrl, queryParamsItem.UserId, queryParamsItem.From, queryParamsItem.To);
                        
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





}
