using DataRecoveryLib.Extensions;
using DataRecoveryLib.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;

namespace DataRecoveryLib
{
    public class PostRequestDataProvider
    {
        FileInfo fileInfo;
        ExcelPackage excelPackage;
        //List<PostRequestParameters> parametersList;

        public PostRequestDataProvider(FileInfo fileInfo)
        {
            this.fileInfo = fileInfo;
            excelPackage = new ExcelPackage(fileInfo);
        }

        public IEnumerable<string> GetCustomerIds(out IEnumerable<string> insertDates)
        {
            using (excelPackage)
            {      
                var ws = excelPackage.Workbook.Worksheets[1];
                var firstRow = ExcelExtension.GetRow(ws, 1);
                var customerIdColumnIndex = GetColumnIndex(firstRow, "customer_id");
                var insertDateColumnIndex = GetColumnIndex(firstRow, "insertdate_customer");

                var customerIdColumn = ExcelExtension.GetColumn(ws, customerIdColumnIndex);
                var customerIds = customerIdColumn.Skip(1);
                insertDates = ExcelExtension.GetColumn(ws, insertDateColumnIndex);
                return customerIds;
            }
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
    }
}
