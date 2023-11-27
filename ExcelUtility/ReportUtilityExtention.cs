using ClosedXML.Excel;
using Samer.ExcelBuilder;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Text;
using static Samer.ExcelBuilder.ExcelTable;

namespace Wellyx.Utility.Reports
{
    public static class ExcelReportUtilityExtention
    {
        public static Stream GenerateExpiredAndTerminatedMembershipsReport(DataSet dataSource)
        {
            if (dataSource.Tables.Count < 3)
                throw new Exception("Dataset must contains tables for actual data, filters and company informaton to present, respectivily.");

            var filters = dataSource.Tables[2];
            var companyInformation = dataSource.Tables[1];
            var actualDataSet = dataSource.Tables[0];
            var calculationDataset = dataSource.Tables[3];

            ExcelTable excelTable = new ExcelTable();
            ExcelTable actualDatasetTable = new ExcelTable();
            excelTable.Link(actualDatasetTable);

            #region build excelTable

            //presenting filters
            for (var rowNumber = 1; rowNumber <= filters.Rows.Count; rowNumber++)
            {
                //dataset row 
                var row = filters.Rows[rowNumber - 1];

                //updating data for filter's row
                ExcelRow excelRow = new ExcelRow(new RowStyle(false), row.ItemArray.Select(x => Convert.ToString(x)).ToArray());

                //spaning accros row, so that filter dataset rows could have enough space to adjust it's charcters.
                excelRow.SpanToMaxRowCells();

                //now checking if it is a first row or value contains date range then it should be bolded.
                if (rowNumber == 1 || excelRow.Values().Any(x => x.StartsWith("From:")))
                {
                    excelRow.RowStyle.FontBold = true;
                }

                excelTable.Rows.Add(excelRow);
            }

            //presenting company information
            for (var rowNumber = 1; rowNumber <= companyInformation.Rows.Count; rowNumber++)
            {
                //dataset row 
                var row = companyInformation.Rows[rowNumber - 1];

                //updating data for filter's row
                ExcelRow excelRow = new ExcelRow(new RowStyle(false), row.ItemArray.Select(x => Convert.ToString(x)).ToArray());

                //spaning accros row, so that filter dataset rows could have enough space to adjust it's charcters.
                excelRow.SpanToMaxRowCells();

                excelTable.Rows.Add(excelRow);
            }

            //added empty row
            excelTable.Rows.Add();

            //now reading columns from actual dataset to present those as a row
            ExcelRow columnsRow = new ExcelRow(new RowStyle() { BackroundColor = XLColor.FromHtml("#6F93AE"), FontColor = XLColor.White, Height = 22, FontBold = true, AlignmentVertical = XLAlignmentVerticalValues.Center });
            var columnIndex = 0;
            foreach (DataColumn column in actualDataSet.Columns)
            {
                if (columnIndex + 1 == actualDataSet.Columns.Count)
                {
                    columnsRow.AddCell(column.ColumnName, new CellStyle { Colspan = 3 });
                }
                else
                {
                    columnsRow.AddCell(column.ColumnName);
                }
                columnIndex++;
            }
            actualDatasetTable.Rows.Add(columnsRow);

            //presenting actual dataset
            for (var rowNumber = 1; rowNumber <= actualDataSet.Rows.Count; rowNumber++)
            {
                //dataset row 
                var row = actualDataSet.Rows[rowNumber - 1];

                var datasetRow = new ExcelRow();

                for (int columnNumber = 0; columnNumber < row.ItemArray.Length; columnNumber++)
                {
                    string columnValue = Convert.ToString(row[columnNumber]);
                    if (columnNumber + 1 == row.ItemArray.Length)
                    {
                        datasetRow.AddCell(columnValue, new CellStyle { Colspan = 3 });
                    }
                    else
                    {
                        datasetRow.AddCell(columnValue);
                    }
                }
                //updating data for actual data's row
                actualDatasetTable.Rows.Add(datasetRow);
            }


            //presenting Calculation Dataset
            var caluculationTable = new ExcelTable();
            caluculationTable.EmptyRowsBeforePresentation = 1;
            caluculationTable.AlignTableEnd = true;

            for (var rowNumber = 1; rowNumber <= calculationDataset.Rows.Count; rowNumber++)
            {
                //dataset row 
                var row = calculationDataset.Rows[rowNumber - 1];
                var cellValues = row.ItemArray.Select(x => Convert.ToString(x));

                var calculationRow = new ExcelRow(new RowStyle(false));
                var isFirstCell = false;
                foreach (var columnValue in cellValues)
                {
                    if (!isFirstCell)
                    {
                        calculationRow.AddCell(columnValue, new CellStyle { Colspan = 1, FontBold = true });
                        isFirstCell = true;
                    }
                    else
                    {
                        calculationRow.AddCell(columnValue, new CellStyle { AlignmentHorizontal = XLAlignmentHorizontalValues.Right });
                    }
                }
                if (rowNumber != 1)
                    calculationRow.RowStyle.TopBorder = XLBorderStyleValues.Thin;

                //updating data for actual data's row
                caluculationTable.Rows.Add(calculationRow);
            }

            #endregion

            //now building report for ExcelTable
            using (var d = ExcelBuilder.Datasets(excelTable, actualDatasetTable, caluculationTable).Build())
            {
                var otherMemebrship = actualDatasetTable.GetColumn("Other Memberships");
                if(otherMemebrship!=null)
                {
                    otherMemebrship.ColumnStyle.Width = 20;
                }
                var totalColumn = caluculationTable.GetColumn("column1");
                if (totalColumn != null)
                {
                    totalColumn.ColumnStyle.Width = 30;
                }
                d.SaveAsFile("hello.xlsx");
                return d.GetStream();
            }
        }
        public static Stream GenerateMemberMembershipsPaymentsReport(DataSet dataSource)
        {
            if (dataSource.Tables.Count < 3)
                throw new Exception("Dataset must contains tables for actual data, filters and branch informaton to present, respectivily.");

            var filters = dataSource.Tables[2];
            var companyInformation = dataSource.Tables[1];
            var actualDataSet = dataSource.Tables[0];

            ExcelTable excelTable = new ExcelTable();

            #region build excelTable

            //presenting filters
            for (var rowNumber = 1; rowNumber <= filters.Rows.Count; rowNumber++)
            {
                //dataset row 
                var row = filters.Rows[rowNumber - 1];

                //updating data for filter's row
                ExcelRow excelRow = new ExcelRow(new RowStyle(false), row.ItemArray.Select(x => Convert.ToString(x)).ToArray());

                //spaning accros row, so that filter dataset rows could have enough space to adjust it's charcters.
                excelRow.SpanToMaxRowCells();

                //now checking if it is a first row or value contains date range then it should be bolded.
                if (rowNumber == 1 || excelRow.Values().Any(x => x.StartsWith("From:")))
                {
                    excelRow.RowStyle.FontBold = true;
                }

                excelTable.Rows.Add(excelRow);
            }

            //presenting company information
            for (var rowNumber = 1; rowNumber <= companyInformation.Rows.Count; rowNumber++)
            {
                //dataset row 
                var row = companyInformation.Rows[rowNumber - 1];

                //updating data for filter's row
                ExcelRow excelRow = new ExcelRow(new RowStyle(false), row.ItemArray.Select(x => Convert.ToString(x).Replace("<br>"," ")).ToArray());

                //spaning accros row, so that filter dataset rows could have enough space to adjust it's charcters.
                excelRow.SpanToMaxRowCells();

                excelTable.Rows.Add(excelRow);
            }

            //added empty row
            excelTable.Rows.Add();

            //now reading columns from actual dataset to present those as a row
            ExcelRow columnsRow = new ExcelRow(new RowStyle() { BackroundColor = XLColor.FromHtml("#6F93AE"), FontColor = XLColor.White, Height = 22, FontBold = true, AlignmentVertical = XLAlignmentVerticalValues.Center });
            foreach (DataColumn column in actualDataSet.Columns)
            {
                columnsRow.AddCell(column.ColumnName);
            }
            excelTable.Rows.Add(columnsRow);

            //presenting actual dataset
            for (var rowNumber = 1; rowNumber <= actualDataSet.Rows.Count; rowNumber++)
            {
                //dataset row 
                var row = actualDataSet.Rows[rowNumber - 1];

                //updating data for actual data's row
                excelTable.Rows.Add(row.ItemArray.Select(x => Convert.ToString(x)).ToArray());
            }

            #endregion

            //now building report for ExcelTable
            using (var d = ExcelBuilder.Datasets(excelTable).Build())
            {
                return d.GetStream();
            }
        }
        public static string GenerateAllMemberMembershipsReport(DataSet dataSource)
        {
            if (dataSource == null || dataSource.Tables[0].Rows.Count == 0)
            {
                return null;
            }

            var details = dataSource.Tables[3];
            var filters = dataSource.Tables[2];
            var companyInformation = dataSource.Tables[1];
            var actualDataSet = dataSource.Tables[0];


            ExcelTable excelTable = new ExcelTable();

            #region build excelTable

            excelTable.Rows.Add(new RowStyle(false) { FontBold = true, }, "Member Memberships").SpanToMaxRowCells();
            excelTable.Rows.Add(new RowStyle(false) { FontBold = false, }, "Search Filter").SpanToMaxRowCells();

            //presenting filters
            for (var rowNumber = 1; rowNumber <= filters.Rows.Count; rowNumber++)
            {
                //dataset row 
                var row = filters.Rows[rowNumber - 1];

                //updating data for filter's row
                ExcelRow excelRow = new ExcelRow(new RowStyle(false), row.ItemArray.Select(x => Convert.ToString(x)).ToArray());

                //spaning accros row, so that filter dataset rows could have enough space to adjust it's charcters.
                excelRow.SpanToMaxRowCells();

                //now checking if it is a first row or value contains date range then it should be bolded.
                if (excelRow.Values().Any(x => x.StartsWith("From:")))
                {
                    excelRow.RowStyle.FontBold = true;
                }

                excelTable.Rows.Add(excelRow);
            }

            //presenting company information
            for (var rowNumber = 1; rowNumber <= companyInformation.Rows.Count; rowNumber++)
            {
                //dataset row 
                var row = companyInformation.Rows[rowNumber - 1];

                //updating data for filter's row
                ExcelRow excelRow = new ExcelRow(new RowStyle(false), row.ItemArray.Select(x => Convert.ToString(x)).ToArray());

                //spaning accros row, so that filter dataset rows could have enough space to adjust it's charcters.
                excelRow.SpanToMaxRowCells();

                excelTable.Rows.Add(excelRow);
            }

            //added empty row
            excelTable.Rows.Add();

            var memberDetails = details.Rows[0].ItemArray.Select(x => Convert.ToString(x)).ToArray();
            var memberdetailsExcelRow = new ExcelRow(new RowStyle(false) { FontBold = true, TopBorder = XLBorderStyleValues.Medium, BottomBorder = XLBorderStyleValues.Medium, AlignmentHorizontal = XLAlignmentHorizontalValues.Center, Height = 40 });
            memberdetailsExcelRow.AddCell(memberDetails[0], new CellStyle { Colspan = 2 });
            memberdetailsExcelRow.AddCell(memberDetails[1], new CellStyle {  BackroundColor = XLColor.FromHtml("#D1D1D1") });
            memberdetailsExcelRow.AddCell(memberDetails[2], new CellStyle { Colspan = 2 });
            memberdetailsExcelRow.AddCell(memberDetails[3], new CellStyle { Colspan = 1 });
            memberdetailsExcelRow.AddCell(memberDetails[4], new CellStyle { Colspan = 3 });
            excelTable.Rows.Add(memberdetailsExcelRow);
            
            //added empty row for space
            excelTable.Rows.Add();
            
            //now reading columns from actual dataset to present those as a row
            ExcelRow columnsRow = new ExcelRow(new RowStyle() { BackroundColor = XLColor.FromHtml("#6F93AE"), FontColor = XLColor.White, Height = 22, FontBold = true, AlignmentVertical = XLAlignmentVerticalValues.Center });
            var columnIndex = 0;
            foreach (DataColumn column in actualDataSet.Columns)
            {
                if (columnIndex + 1 == actualDataSet.Columns.Count)
                {
                    columnsRow.AddCell(column.ColumnName, new CellStyle { Colspan = 3 });
                }
                else
                {
                    columnsRow.AddCell(column.ColumnName);
                }
                columnIndex++;
            }
            excelTable.Rows.Add(columnsRow);

            //presenting actual dataset
            for (var rowNumber = 1; rowNumber <= actualDataSet.Rows.Count; rowNumber++)
            {
                //dataset row 
                var row = actualDataSet.Rows[rowNumber - 1];

                var datasetRow = new ExcelRow();

                for (int columnNumber = 0; columnNumber < row.ItemArray.Length; columnNumber++)
                {
                    string columnValue = Convert.ToString(row[columnNumber]);
                    if (columnNumber == 0)
                    {
                        columnValue = rowNumber.ToString();
                    }
                    if (columnNumber + 1 == row.ItemArray.Length)
                    {
                        datasetRow.AddCell(columnValue, new CellStyle { Colspan = 3 });
                    }
                    else
                    {
                        datasetRow.AddCell(columnValue);
                    }
                }
                //updating data for actual data's row
                excelTable.Rows.Add(datasetRow);
            }


            #endregion

            //now building report for ExcelTable
            using (var d = ExcelBuilder.Datasets(excelTable).Build())
            {
                d.SaveAsFile("Hello1.xlsx");
                return "";
            }



        }
    }
}
