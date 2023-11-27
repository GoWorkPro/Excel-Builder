
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Math;
using Microsoft.ML;
using Microsoft.ML.Data;
using Samer.CsvBuilder;
using System.Data;
using Wellyx.Utility.ExcelReports;
using static Program;

class Program
{

    static void Main(string[] args)
    {
        #region test data
        // #region datasets
        DataTable dataTableFilter = new DataTable("SampleDataFilters");
        //Add the column names
         dataTableFilter.Columns.Add("column1", typeof(string));

        dataTableFilter.Rows.Add("All Customers");
        dataTableFilter.Rows.Add("Search Filter");
        dataTableFilter.Rows.Add("Customer Type: All");
        dataTableFilter.Rows.Add("Customer Status: Active");
        dataTableFilter.Rows.Add("From: 03 - 01 - 2022 To: 03 - 11 - 2023");

        DataTable companyInfo = new DataTable("company");
        //Add the column names
         companyInfo.Columns.Add("column1", typeof(string));
        companyInfo.Rows.Add("London");
        companyInfo.Rows.Add("Add01a, City");
        companyInfo.Rows.Add("London, sef83n");
        companyInfo.Rows.Add("T: +441234567890");
        companyInfo.Rows.Add("E:");



        DataTable actualDataset = new DataTable("SampleDataActual");

        actualDataset.Columns.Add("Sr. <br /> ", typeof(int));
        actualDataset.Columns.Add("Branch", typeof(string));
        actualDataset.Columns.Add("Branch Code", typeof(int));
        actualDataset.Columns.Add("Name", typeof(string));
        actualDataset.Columns.Add("Email", typeof(string));
        actualDataset.Columns.Add("Mobile <br />", typeof(long));
        actualDataset.Columns.Add("Membership Created Date", typeof(string));

        actualDataset.Rows.Add(1, "Branch 1", 82, "Alex Will", "Alw@yopmail.com", 44578963219, "18.02.2022 - 11:00AM");
        actualDataset.Rows.Add(2, "Branch 1", 82, "eqw Will <br />", "eqw@yopmail.com", 4324, "18.02.2022 - 11:00AM");
        actualDataset.Rows.Add(3, "Branch 1", 82, "eqw Will", "32@yopmail.com <br />", 4455353478963219, "18.02.2022 - 11:00AM");




        // DataTable resultData = new DataTable("resultDataset");

        // Add the column names
        // resultData.Columns.Add("col", typeof(string));
        // resultData.Columns.Add("no", typeof(string));
        // Sample metadata
        // resultData.Rows.Add("Totals");
        // resultData.Rows.Add("Gross Total:", "$200");
        // resultData.Rows.Add("Total:", "$190");

        // DataTable customerDetails = new DataTable("customerDetails");
        // Sample metadata
        // customerDetails.Columns.Add("col1", typeof(string));
        // customerDetails.Columns.Add("col2", typeof(string));
        // customerDetails.Columns.Add("col3", typeof(string));
        // customerDetails.Columns.Add("col4", typeof(string));
        // customerDetails.Columns.Add("col5", typeof(string));
        // customerDetails.Rows.Add(new string[] { "sc 01", "Member", "E: test.test1@gmail.com", "M: +145455454", "A: bananas, London, Nanakbznzn, Afghanistan, iqi862" });
        // #endregion

        // #region Simple CustomerReport
        // ExcelTable simpleexcelTable = new ExcelTable();



        // simpleexcelTable.Rows.Add("Sr", "Name", "Gender", "Customer Type", "Email", "Mobile", "Source", "Occupation", "Date of Birth", "Date Created");
        // simpleexcelTable.Rows.Add("1", "cfgewdfcrweg gfgfgfg", "Not Specified", "Client", "test@yopmail.com", "", "Wellyx-Core", "", "09-08-2023", "");
        // simpleexcelTable.Rows.Add("2", "hsn test", "Not Specified", "Member", "hsntest@gmail.com", "", "Wellyx-Core", "", "17-10-2022", "");

        // #endregion

        // #region CustomerReport
        // ExcelTable excelTable = new ExcelTable();

        // Add rows based on the provided lines of text
        //var rowABC = new ExcelRow(new RowStyle(false) { FontBold = true, FontSize = 12 }, "All Customers");
        // excelTable.Rows.Add(rowABC);
        // excelTable.Rows.Add(new RowStyle(false), "Search Filter").SpanToMaxRowCells();
        // excelTable.Rows.Add(new RowStyle(false), "Customer Type: All").SpanToMaxRowCells();
        // excelTable.Rows.Add(new RowStyle(false), "Customer Status: Active").SpanToMaxRowCells();
        // excelTable.Rows.Add(new RowStyle(false) { FontBold = true, FontSize = 13 }, "From: 03-01-2022 To: 03-11-2023").SpanToMaxRowCells();
        // excelTable.Rows.Add(new RowStyle(false), "London").SpanToMaxRowCells();
        // excelTable.Rows.Add(new RowStyle(false), "Add01a, City").SpanToMaxRowCells();
        // excelTable.Rows.Add(new RowStyle(false), "London, sef83n").SpanToMaxRowCells();
        // excelTable.Rows.Add(new RowStyle(false), "T: +441234567890").SpanToMaxRowCells();
        // excelTable.Rows.Add(new RowStyle(false), "E: ").SpanToMaxRowCells();

        // excelTable.Rows.Add().SpanToMaxRowCells();

        // excelTable.Rows.Add(new RowStyle
        // {
        //     FontBold = true,
        //     Height = 25,
        //     AlignmentVertical = XLAlignmentVerticalValues.Center,
        //     BackroundColor = XLColor.FromHtml("#6F93AE"),
        //     FontColor = XLColor.White
        // }, "Sr", "Name", "Gender", "Customer Type", "Email", "Mobile", "Source", "Occupation", "Date of Birth", "Date Created").Freeze();
        // excelTable.Rows.Add("1", "cfgewdfcrweg gfgfgfg", "Not Specified", "Client", "sdfsddffsd@yopmail.com", "", "Wellyx-Core", "", "09-08-2023", "");
        // excelTable.Rows.Add("2", "hsn test", "Not Specified", "Member", "hsntest@gmail.com", "", "Wellyx-Core", "", "17-10-2022", "");

        // #endregion

        // #region MailingListReport
        // ExcelTable excelTable1 = new ExcelTable();
        // ExcelTable mailingReportDataTable = new ExcelTable();
        // excelTable1.Link(mailingReportDataTable);

        // excelTable1.Rows.Add(new RowStyle(false) { FontBold = true, FontSize = 13 }, "Mailing List").SpanToMaxRowCells();
        // excelTable1.Rows.Add(new RowStyle(false), "Search Filter").SpanToMaxRowCells();
        // excelTable1.Rows.Add(new RowStyle(false), "Customer Type: All").SpanToMaxRowCells();
        // excelTable1.Rows.Add(new RowStyle(false), "Customer Status: Active").SpanToMaxRowCells();
        // excelTable1.Rows.Add(new RowStyle(false) { FontBold = true }, "From: 03-01-2022 To: 03-11-2023").SpanToMaxRowCells();
        // excelTable1.Rows.Add(new RowStyle(false), "London").SpanToMaxRowCells();
        // excelTable1.Rows.Add(new RowStyle(false), "Add01a, City").SpanToMaxRowCells();
        // excelTable1.Rows.Add(new RowStyle(false), "London, sef83n").SpanToMaxRowCells();
        // excelTable1.Rows.Add(new RowStyle(false), "T: +441234567890").SpanToMaxRowCells();
        // excelTable1.Rows.Add(new RowStyle(false), "E: ").SpanToMaxRowCells();

        // excelTable1.Rows.Add().SpanToMaxRowCells();


        // var topRow = new ExcelRow(new RowStyle
        // {
        //     FontBold = true,
        //     TopBorder = XLBorderStyleValues.Double,
        //     Height = 25,
        //     AlignmentVertical = XLAlignmentVerticalValues.Center,
        //     FontSize = 12
        // });
        // topRow.AddCell("Personal Information", new CellStyle
        // {
        //     Colspan = 6,
        //     AlignmentHorizontal = XLAlignmentHorizontalValues.Center,
        //     BackroundColor = XLColor.Gray,
        //     FontColor = XLColor.White,
        // });

        // topRow.AddCell("Address Informaton", new CellStyle
        // {
        //     AlignmentHorizontal = XLAlignmentHorizontalValues.Center,
        //     BackroundColor = XLColor.BlueBell,
        //     FontColor = XLColor.White,
        //     Colspan = 5
        // });
        // topRow.AddCell("Permissions", new CellStyle
        // {
        //     AlignmentHorizontal = XLAlignmentHorizontalValues.Center,
        //     BackroundColor = XLColor.BlueGray,
        //     FontColor = XLColor.White,
        //     Colspan = 5,
        // });
        // mailingReportDataTable.Rows.Add(topRow);

        // mailingReportDataTable.Rows.Add(new RowStyle { FontBold = true, AlignmentVertical = XLAlignmentVerticalValues.Center }, "Branch", "Name", "Email", "Mobile", "Membership Created Date", "Membership", "Status", "Start Date", "End Date", "Roll Over", "Reason", "Notes", "Expired/Terminated Date");
        // mailingReportDataTable.Rows.Add("Branch 1", "Babar Hassan", "babar@yopmail.com", "+442071773406", "13-03-2023<br>06:02 AM", "AppleDiscount33", "Terminated", "13-03-2023", "13-03-2273", "No", "Reason not given", "N/A", "26-09-2023<br>02:46 PM");
        // mailingReportDataTable.Rows.Add("Branch 1", "Babar Hassan", "babarh@yopmail.com", "+44 9797 67", "24-03-2023<br>06:19 AM", "5 Day Membership - 3", "Terminated", "24-03-2023", "28-03-2023", "No", "Reason not given", "N/A", "24-03-2023<br>06:20 AM");
        // mailingReportDataTable.Rows.Add("Branch 1", "Babar Hassan", "babarh@yopmail.com", "+44 9797 67", "23-03-2023<br>04:48 AM", "5 Day Membership - 2", "Terminated", "23-03-2023", "27-03-2023", "No", "Reason not given", "N/A", "24-03-2023<br>06:17 AM");
        // mailingReportDataTable.Rows.Add("Branch 1", "Babar Hassan", "babarh@yopmail.com", "+44 9797 67", "22-03-2023<br>05:05 PM", "5 Day Membership - 1", "Terminated", "22-03-2023", "26-03-2023", "No", "Reason not given", "N/A", "23-03-2023<br>04:47 AM");
        // mailingReportDataTable.Rows.Add("Branch 1", "Babar Hassan", "babarh@yopmail.com", "+44 9797 67", "22-03-2023<br>04:50 PM", "5 Day Membership", "Terminated", "22-03-2023", "26-03-2023", "No", "Reason not given", "N/A", "22-03-2023<br>04:57 PM");
        // mailingReportDataTable.Rows.Add("Branch 1", "Sy Mehr", "sqmehr@yopmail.com", "+92 3345 045595", "26-04-2022<br>09:24 AM", "Crossfit Membership - Copy 2", "Terminated", "26-04-2022", "26-04-2023", "Yes", "Reason not given", "N/A", "15-03-2023<br>10:30 AM");
        // mailingReportDataTable.Rows.Add("Branch 1", "Sy Mehr", "sqmehr@yopmail.com", "+92 3345 045595", "14-03-2023<br>01:03 PM", "Yearly membership", "Terminated", "14-03-2023", "14-03-2028", "No", "Reason not given", "N/A", "15-03-2023<br>10:28 AM");
        // mailingReportDataTable.Rows.Add("Branch 1", "Babar Hassan", "babar@yopmail.com", "+442071773406", "13-03-2023<br>06:07 AM", "Crossfit Membership - 2", "Terminated", "13-03-2023", "13-03-2024", "Yes", "Reason not given", "N/A", "13-03-2023<br>06:53 AM");
        // mailingReportDataTable.Rows.Add("Branch 1", "Babar Hassan", "babar@yopmail.com", "+442071773406", "13-03-2023<br>05:45 AM", "Crossfit Membership - 1", "Terminated", "13-03-2023", "13-03-2024", "Yes", "Reason not given", "N/A", "13-03-2023<br>06:04 AM");
        // mailingReportDataTable.Rows.Add("Branch 1", "Babar Hassan", "babar@yopmail.com", "+442071773406", "10-03-2023<br>01:30 PM", "WG - 12 Months x 4 $59.99", "Terminated", "10-03-2023", "06-04-2023", "Yes", "Reason not given", "N/A", "10-03-2023<br>01:32 PM");
        // mailingReportDataTable.Rows.Add("London", "shani 001", "shani66@yopmail.com", "N/A", "10-03-2023<br>10:25 AM", "interval", "Terminated", "10-03-2023", "10-03-2024", "No", "Reason not given", "N/A", "10-03-2023<br>12:15 PM");
        // mailingReportDataTable.Rows.Add("Branch 1", "Babar Hassan", "babar@yopmail.com", "+442071773406", "10-03-2023<br>11:53 AM", "Crossfit Membership", "Terminated", "10-03-2023", "10-03-2024", "Yes", "Reason not given", "N/A", "10-03-2023<br>11:59 AM");
        // mailingReportDataTable.Rows.Add("Branch 1", "Babar Hassan", "babar@yopmail.com", "+442071773406", "28-02-2023<br>11:52 AM", "Yearly membership", "Terminated", "15-03-2023", "15-03-2025", "No", "Reason not given", "N/A", "01-03-2023<br>05:56 AM");
        // mailingReportDataTable.Rows.Add("London Branch", "Ham Dard", "hamdard@yopmail.com", "N/A", "06-01-2023<br>11:28 AM", "12 month winter package - 1", "Terminated", "06-01-2023", "06-01-2024", "Yes", "Reason not given", "N/A", "06-01-2023<br>01:34 PM");
        // mailingReportDataTable.Rows.Add("London Branch", "Ham Dard", "hamdard@yopmail.com", "N/A", "05-01-2023<br>01:45 PM", "12 month winter package", "Terminated", "05-01-2023", "05-01-2024", "Yes", "Reason not given", "N/A", "06-01-2023<br>11:11 AM");


        // Add more rows in a similar way
        // #endregion

        // #region AllSalesDetailByCustomer
        //ExcelTable excelTable2 = new ExcelTable();

        // Add rows based on the provided lines of text
        // excelTable2.Rows.Add(new RowStyle(false) { FontBold = true, FontSize = 12 }, "All Sales Detail By Customer").SpanToMaxRowCells();
        // excelTable2.Rows.Add(new RowStyle(false), "Search Filter").SpanToMaxRowCells();
        // excelTable2.Rows.Add(new RowStyle(false), "Customer Type: All").SpanToMaxRowCells();
        // excelTable2.Rows.Add(new RowStyle(false), "Customer Status: Active").SpanToMaxRowCells();
        // excelTable2.Rows.Add(new RowStyle(false) { FontBold = true, FontSize = 13 }, "From: 03-01-2022 To: 03-11-2023").SpanToMaxRowCells();
        // excelTable2.Rows.Add(new RowStyle(false), "London").SpanToMaxRowCells();
        // excelTable2.Rows.Add(new RowStyle(false), "Add01a, City").SpanToMaxRowCells();
        // excelTable2.Rows.Add(new RowStyle(false), "London, sef83n").SpanToMaxRowCells();
        // excelTable2.Rows.Add(new RowStyle(false), "T: +441234567890").SpanToMaxRowCells();
        // excelTable2.Rows.Add(new RowStyle(false), "E: ").SpanToMaxRowCells();

        // excelTable2.Rows.Add().SpanToMaxRowCells();

        // excelTable2.Rows.Add(new ExcelRow(new RowStyle
        // {
        //     FontBold = true,
        //     Height = 25,
        //     AlignmentVertical = XLAlignmentVerticalValues.Center,
        //     TopBorder = XLBorderStyleValues.Double
        // }, "Sr", "Name", "Gender", "Customer Type", "Email", "Mobile", "Source", "Occupation", "Date of Birth", "Date Created"));
        // excelTable2.Rows.Add(new ExcelRow("1", "cfgewdfcrweg gfgfgfg", "Not Specified", "Client", "sdfsddffsd@yopmail.com", "", "Wellyx-Core", "", "09-08-2023", ""));
        // excelTable2.Rows.Add(new ExcelRow("2", "hsn test", "Not Specified", "Member", "hsntest@gmail.com", "", "Wellyx-Core", "", "17-10-2022", ""));


        // ExcelTable excelTable3 = new ExcelTable();
        // excelTable3.AlignTableEnd = true;
        // excelTable3.EmptyRowsBeforePresentation = 2;
        // excelTable3.Rows.Add("Gross Total", "$0.00");
        // excelTable3.Rows.Add("Total Benefits Used Discounts", "$0.00");
        // excelTable3.Rows.Add("Total line-item Discounts", "$0.00");
        // excelTable3.Rows.Add(new RowStyle { TopBorder = XLBorderStyleValues.Double }, "Net Total", "$0.00");


        // ExcelTable excelTable4 = new ExcelTable();
        // excelTable4.EmptyRowsBeforePresentation = 2;

        // excelTable4.Rows.Add(new RowStyle(false), new CellStyle { FontBold = true, Colspan = 4 }, "Calculation Formulas:");
        // excelTable4.Rows.Add(new RowStyle(false), new CellStyle { Colspan = 6 }, "Gross Total = (Per Unit Price * Total Quantity) + Service Charges + Taxes + Tips.");
        // excelTable4.Rows.Add(new RowStyle(false), new CellStyle { Colspan = 6 }, "Net Total = Gross Total - Total Discounted Amount - Service Charges - Taxes - Tips.");


        // #endregion

        // #region Payment Summary
        // ExcelTable excelPaymentSummary = new ExcelTable();
        // excelPaymentSummary.Rows.Add(new RowStyle(false) { FontBold = true }, "Payments Summary").SpanToMaxRowCells();
        // excelPaymentSummary.Rows.Add(new RowStyle(false), "Search Filter").SpanToMaxRowCells();
        // excelPaymentSummary.Rows.Add(new RowStyle(false), "From:01-11-2023 To:01-11-2023").SpanToMaxRowCells();
        // excelPaymentSummary.Rows.Add().SpanToMaxRowCells();

        // ExcelRow excelRow = new ExcelRow(new RowStyle { BackroundColor = XLColor.FromHtml("#6F93AE"), FontColor = XLColor.White, Height = 25, AlignmentVertical = XLAlignmentVerticalValues.Center });
        // excelRow.AddCell("Payment Method");
        // excelRow.AddCell("Payment Gateway");
        // excelRow.AddCell("Total Amount", new CellStyle { AlignmentHorizontal = XLAlignmentHorizontalValues.Right, Colspan = 3 });
        // excelPaymentSummary.Rows.Add(excelRow);

        // ExcelRow excelRow1 = new ExcelRow();
        // excelRow1.AddCell("Cash", new CellStyle { Rowspan = 4, AlignmentVertical = XLAlignmentVerticalValues.Top, FontBold = true });
        // excelRow1.AddCell("", new CellStyle { Colspan = 4 });
        // excelPaymentSummary.Rows.Add(excelRow1);

        // ExcelRow excelRow2 = new ExcelRow();
        // excelRow2.AddCell();
        // excelRow2.AddCell("Cash");
        // excelRow2.AddCell("$0.00", new CellStyle { AlignmentHorizontal = XLAlignmentHorizontalValues.Right, Colspan = 3 });
        // excelPaymentSummary.Rows.Add(excelRow2);

        // ExcelRow excelRow3 = new ExcelRow();
        // excelRow3.AddCell();
        // excelRow3.AddCell("Cash Refunds");
        // excelRow3.AddCell("($0.00)", new CellStyle { AlignmentHorizontal = XLAlignmentHorizontalValues.Right, Colspan = 3 });
        // excelPaymentSummary.Rows.Add(excelRow3);


        // excelPaymentSummary.Rows.Add(new ExcelCell(), new ExcelCell("Cash Grand Totals", new CellStyle { FontBold = true }), new ExcelCell("($0.00)", new CellStyle { AlignmentHorizontal = XLAlignmentHorizontalValues.Right, Colspan = 3 }));

        // excelPaymentSummary.Rows.Add(new ExcelCell("3rd Party", new CellStyle { Rowspan = 3, AlignmentVertical = XLAlignmentVerticalValues.Top, FontBold = true }),
        //      new ExcelCell("", new CellStyle { Colspan = 4 }));

        // excelPaymentSummary.Rows.Add(new ExcelCell(), new ExcelCell("Others"), new ExcelCell("$0.00", new CellStyle { AlignmentHorizontal = XLAlignmentHorizontalValues.Right, Colspan = 3 }));
        // excelPaymentSummary.Rows.Add(new ExcelCell(), new ExcelCell("3rd Party Grand Total", new CellStyle { FontBold = true }), new ExcelCell("$0.00", new CellStyle { AlignmentHorizontal = XLAlignmentHorizontalValues.Right, Colspan = 3 }));


        // ExcelTable excelPaymentSummary1 = new ExcelTable();
        // excelPaymentSummary1.AlignTableEnd = true;
        // excelPaymentSummary1.EmptyRowsBeforePresentation = 2;
        // excelPaymentSummary1.Rows.Add(new RowStyle(false), new ExcelCell("Total Payments"), new ExcelCell("$0.00", new CellStyle { AlignmentHorizontal = XLAlignmentHorizontalValues.Right, FontBold = true }));
        // excelPaymentSummary1.Rows.Add(new RowStyle(false) { TopBorder = XLBorderStyleValues.Double }, new ExcelCell("Total Refunds"), new ExcelCell("$0.00", new CellStyle { AlignmentHorizontal = XLAlignmentHorizontalValues.Right, FontBold = true }));
        // excelPaymentSummary1.Rows.Add(new RowStyle(false) { FontBold = true }, new ExcelCell("Grand Total:"), new ExcelCell("($0.00)", new CellStyle { AlignmentHorizontal = XLAlignmentHorizontalValues.Right, FontBold = true }));

        // ExcelTable excelPaymentSummary2 = new ExcelTable();
        // excelPaymentSummary2.EmptyRowsBeforePresentation = 2;
        // excelPaymentSummary2.Rows.Add(new RowStyle(false), new ExcelCell("Calculation Formulas :", new CellStyle() { Colspan = 2, FontBold = true }));

        // var excelRichCell = new ExcelCell(new CellStyle() { Colspan = 2 });
        // excelRichCell.RichText.Add(new RichTextValue { Value = "Grand Total", Bold = true });
        // excelRichCell.RichText.AddALine(new RichTextValue { Value = " = Total Payments - Total Refunds" });
        // excelPaymentSummary2.Rows.Add(new RowStyle(false), excelRichCell);
        // excelPaymentSummary2.Rows.Add(new RowStyle(false), new ExcelCell("All payments are inclusive of taxes, tips and service charges.", new CellStyle() { Colspan = 2, BackroundColor = XLColor.FromHtml("#FFD700") }));


        // #endregion



        // using (var d = ExcelBuilder.Datasets(simpleexcelTable).Build())
        // {
        //     // Specify the file path where you want to save the Excel file
        //     string filePath = "DataTableToExcel-1-Simple-Customer-Report.xlsx";
        //     d.SaveAsFile(filePath);

        // }

        // using (var d = ExcelBuilder.Datasets(excelTable).Build())
        // {
        //     // Specify the file path where you want to save the Excel file
        //     string filePath = "DataTableToExcel-2-Customer-Report.xlsx";
        //     d.SaveAsFile(filePath);
        // }

        // using (var d = ExcelBuilder.Datasets(excelTable1, excelTableMailingReport).Build())
        // {
        //     var column = excelTableMailingReport.GetColumn("Personal Information");

        //     excelTableMailingReport.GetColumn(column.EndColumnNumber).ColumnStyle.RightBorder = XLBorderStyleValues.Double;
        //     // Specify the file path where you want to save the Excel file
        //     string filePath = "DataTableToExcel-3-Mailing-List-Report.xlsx";
        //     d.SaveAsFile(filePath);
        // }

        // using (var d = ExcelBuilder.Datasets(excelTable2, excelTable3, excelTable4).Build())
        // {
        //     // Specify the file path where you want to save the Excel file
        //     string filePath = "DataTableToExcel-4-All-Sales-Detail-By-Customer-Report.xlsx";
        //     d.SaveAsFile(filePath);
        // }

        // //Save the Excel package to a file
        // using (var d = ExcelBuilder.Datasets(excelPaymentSummary, excelPaymentSummary1, excelPaymentSummary2).Build())
        // {
        //     // Specify the file path where you want to save the Excel file
        //     string filePath = "DataTableToExcel-5-Payment-Summary.xlsx";
        //     d.SaveAsFile(filePath);
        // }

        // using (var d = ExcelBuilder.Datasets(new Worksheet("SimpleReport", simpleexcelTable), new Worksheet("allCustomerReport", excelTable), new Worksheet("mailingListReport", mailingReportDataTable), new Worksheet("AllSalesDetailReport", excelTable2, excelTable3, excelTable4), new Worksheet("Payment-Summary", excelPaymentSummary, excelPaymentSummary1, excelPaymentSummary2)).Build())
        // {
        //     // Specify the file path where you want to save the Excel file
        //     string filePath = "composit-report.xlsx";
        //     excelPaymentSummary1.GetColumn("column1").ColumnStyle.BackroundColor = XLColor.Blue;
        //     d.SaveAsFile(filePath);
        // }


        DataSet dataSet = new DataSet();
        dataSet.Tables.Add(actualDataset);
        dataSet.Tables.Add(companyInfo);
        dataSet.Tables.Add(dataTableFilter);
        //Wellyx.Utility.Reports.ExcelReportUtilityExtention.GenerateAllMemberMembershipsReport(dataSet);

        var stream = CsvReportUtilityExtention.GenerateExpiredAndTerminatedMembershipsReport(dataSet);
        using (FileStream fileStream = File.Create("hello CSV.csv"))
        {
            stream.CopyTo(fileStream);
        }

        #endregion test data

        //var context = new MLContext();

        //// Load your data
        //var dataView = context.Data.LoadFromTextFile<ReviewData>("hotel-reviews.txt", separatorChar: ',', hasHeader: true);

        //// Tokenization and vectorization using TextFeaturizing
        //var pipeline = context.Transforms.Text.FeaturizeText("Features", "Review");

        //var transformedData = pipeline.Fit(dataView).Transform(dataView);

        //// Convert IDataView to IEnumerable for in-memory processing
        //var reviewData = context.Data.CreateEnumerable<ReviewData>(transformedData, reuseRowObject: false).ToList();

        //// User query
        //var userQuery = "The hotel's shuttle service to nearby attractions is a thoughtful convenience for guests.";

        //// Vectorize user query
        //var vectorizer = pipeline.Fit(context.Data.LoadFromEnumerable(new List<ReviewData> { new ReviewData { Review = userQuery } }));
        //var userQueryVector = vectorizer.Transform(context.Data.LoadFromEnumerable(new List<ReviewData> { new ReviewData { Review = userQuery } }));

        //// Extract the Features column for the user query
        //var userQueryArray = context.Data.CreateEnumerable<ReviewData>(userQueryVector, reuseRowObject: false).First().Features;

        //// Calculate similarity scores using cosine similarity
        //var similarityScores = reviewData.Select(review => new
        //{
        //    Review = review.Review,
        //    SimilarityScore = EuclideanDistance(userQueryArray, review.Features)
        //});

        //// Sort by similarity score in descending order
        //var sortedSimilarityScores = similarityScores.OrderByDescending(item => item.SimilarityScore);

        //// Display the sorted reviews
        //foreach (var item in sortedSimilarityScores)
        //{
        //    Console.WriteLine($"Review: {item.Review}, Similarity Score: {item.SimilarityScore}");
        //}
    }

    static float EuclideanDistance(float[] vector1, float[] vector2)
    {
        return (float)Math.Sqrt(vector1.Zip(vector2, (a, b) => Math.Pow(a - b, 2)).Sum());
    }
    static string PreprocessText(string text)
    {
        // Convert to lowercase
        text = text.ToLower();

        // Remove punctuation
        text = new string(text.Where(c => !char.IsPunctuation(c)).ToArray());

        // Add additional text preprocessing steps as needed

        return text;
    }

    public class ReviewData
    {
        [LoadColumn(0)]
        public string Review;
        [LoadColumn(2, 51), VectorType(50)]
        public float[] Features;
    }

}
