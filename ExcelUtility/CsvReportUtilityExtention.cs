using Accord.Diagnostics;
using Samer.CsvBuilder;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;

namespace Wellyx.Utility.ExcelReports
{
    public static class CsvReportUtilityExtention
    {
        public static Stream GenerateExpiredAndTerminatedMembershipsReport(DataSet dataSource)
        {
            var filters = dataSource.Tables[2];
            var companyInformation = dataSource.Tables[1];
            var actualDataSet = dataSource.Tables[0];

            var builder = CsvBuilder.Datasets(filters, companyInformation, actualDataSet);
            builder.ValueRenderEvent += (string value, Samer.CsvBuilder.ValueType type, int column, int row, int tableIndex, int actualRow) =>
            {
                if (tableIndex == 2 && row >= 2 && column == 1)
                {
                    return (row - 1).ToString();
                }
                return value;
            };
            return builder.Build(2).GetStream();
        }
        public static Stream GenerateMemberMembershipsPaymentsReport(DataSet dataSource)
        {
            var filters = dataSource.Tables[2];
            var companyInformation = dataSource.Tables[1];
            var actualDataSet = dataSource.Tables[0];
            var details = dataSource.Tables[3];
            ICsvBuilder builder = CsvBuilder.Datasets(filters, companyInformation, actualDataSet);

            return builder.Build(2).GetStream();
        }

        static string func(string value, Samer.CsvBuilder.ValueType valueType)
        {
            return value;
        }
    }
}
