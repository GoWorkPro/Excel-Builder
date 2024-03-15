// Create two ExcelTable instances: excelTable1 and mailingReportDataTable
using ClosedXML.Excel;
using GoWorkPro.ExcelBuilder;
using System.Drawing;
using static GoWorkPro.ExcelBuilder.ExcelTable;

ExcelTable excelTable1 = new ExcelTable();
ExcelTable mailingReportDataTable = new ExcelTable();

// Link excelTable1 to mailingReportDataTable
excelTable1.Link(mailingReportDataTable);

// Add header and filter information to excelTable1
excelTable1.Rows.Add(new RowStyle(false) { FontBold = true, FontSize = 13 }, "Mailing List").SpanToMaxRowCells();
excelTable1.Rows.Add(new RowStyle(false), "Search Filter").SpanToMaxRowCells();
excelTable1.Rows.Add(new RowStyle(false), "Customer Type: All").SpanToMaxRowCells();
excelTable1.Rows.Add(new RowStyle(false), "Customer Status: Active").SpanToMaxRowCells();
excelTable1.Rows.Add(new RowStyle(false) { FontBold = true }, "From: 03-01-2022 To: 03-11-2023").SpanToMaxRowCells();
excelTable1.Rows.Add(new RowStyle(false), "London").SpanToMaxRowCells();
excelTable1.Rows.Add(new RowStyle(false), "Add01a, City").SpanToMaxRowCells();
excelTable1.Rows.Add(new RowStyle(false), "London, sef83n").SpanToMaxRowCells();
excelTable1.Rows.Add(new RowStyle(false), "T: +441234567890").SpanToMaxRowCells();
excelTable1.Rows.Add(new RowStyle(false), "E: ").SpanToMaxRowCells();

excelTable1.Rows.Add().SpanToMaxRowCells();

// Create a top row for mailingReportDataTable with multi-row header and distinct styling for different sections
var topRow = new ExcelRow(new RowStyle
{
    FontBold = true,
    TopBorder = XLBorderStyleValues.Double,
    Height = 25,
    AlignmentVertical = XLAlignmentVerticalValues.Center,
    FontSize = 12
});
topRow.AddCell("Customer Enrollment Information", new CellStyle
{
    Colspan = 5,
    AlignmentHorizontal = XLAlignmentHorizontalValues.Center,
    BackroundColor = XLColor.Gray,
    FontColor = XLColor.White,
});

topRow.AddCell("Balance", new CellStyle
{
    AlignmentHorizontal = XLAlignmentHorizontalValues.Center,
    BackroundColor = XLColor.BlueBell,
    FontColor = XLColor.White,
    Colspan = 2
});
topRow.AddCell("Earned", new CellStyle
{
    AlignmentHorizontal = XLAlignmentHorizontalValues.Center,
    BackroundColor = XLColor.BlueGray,
    FontColor = XLColor.White,
    Colspan = 2,
});
topRow.AddCell("Redeemed", new CellStyle
{
    AlignmentHorizontal = XLAlignmentHorizontalValues.Center,
    BackroundColor = XLColor.BlueGray,
    FontColor = XLColor.White,
    Colspan = 2,
}); topRow.AddCell("Expired", new CellStyle
{
    AlignmentHorizontal = XLAlignmentHorizontalValues.Center,
    BackroundColor = XLColor.BlueGray,
    FontColor = XLColor.White,
    Colspan = 2,
});
mailingReportDataTable.Rows.Add(topRow);

// Add column headers to mailingReportDataTable
mailingReportDataTable.Rows.Add(new RowStyle { FontBold = true, AlignmentVertical = XLAlignmentVerticalValues.Center }, "Branch", "Name", "Email", "Mobile", "Membership Created Date", "Membership", "Status", "Start Date", "End Date", "Roll Over", "Reason", "Notes", "Expired/Terminated Date");

// Add data rows to mailingReportDataTable
mailingReportDataTable.Rows.Add("Branch 1", "Babar Hassan", "babar@yopmail.com", "+442071773406", "13-03-2023 06:02 AM", "AppleDiscount33", "Terminated", "13-03-2023", "13-03-2273", "No", "Reason not given", "N/A", "26-09-2023 02:46 PM");
mailingReportDataTable.Rows.Add("Branch 1", "Sy Mehr", "sqmehr@yopmail.com", "+92 3345 045595", "26-04-2022 09:24 AM", "Crossfit Membership - Copy 2", "Terminated", "26-04-2022", "26-04-2023", "Yes", "Reason not given", "N/A", "15-03-2023 10:30 AM");
mailingReportDataTable.Rows.Add("Branch 1", "Babar Hassan", "babar@yopmail.com", "+442071773406", "13-03-2023 06:07 AM", "Crossfit Membership - 2", "Terminated", "13-03-2023", "13-03-2024", "Yes", "Reason not given", "N/A", "13-03-2023 06:53 AM");
mailingReportDataTable.Rows.Add("Branch 1", "Babar Hassan", "babar@yopmail.com", "+442071773406", "10-03-2023 01:30 PM", "WG - 12 Months x 4 $59.99", "Terminated", "10-03-2023", "06-04-2023", "Yes", "Reason not given", "N/A", "10-03-2023 01:32 PM");
mailingReportDataTable.Rows.Add("London", "shani 001", "shani66@yopmail.com", "N/A", "10-03-2023 10:25 AM", "interval", "Terminated", "10-03-2023", "10-03-2024", "No", "Reason not given", "N/A", "10-03-2023 12:15 PM");
mailingReportDataTable.Rows.Add("Branch 1", "Babar Hassan", "babar@yopmail.com", "+442071773406", "10-03-2023 11:53 AM", "Crossfit Membership", "Terminated", "10-03-2023", "10-03-2024", "Yes", "Reason not given", "N/A", "10-03-2023 11:59 AM");
mailingReportDataTable.Rows.Add("London Branch", "Ham Dard", "hamdard@yopmail.com", "N/A", "06-01-2023 11:28 AM", "12 month winter package - 1", "Terminated", "06-01-2023", "06-01-2024", "Yes", "Reason not given", "N/A", "06-01-2023 01:34 PM");

var rightAlignedTable = new ExcelTable() { AlignTableEnd = true };


var excelRowTableEnd = new ExcelRow();

rightAlignedTable.Rows.Add(new ExcelCell("excelRowTableEnd", new CellStyle
{
    Colspan = 3
}), new ExcelCell("excelRowTableEnd1", new CellStyle { FontBold = true }));

using (var d = ExcelBuilder.Datasets(excelTable1, mailingReportDataTable, rightAlignedTable).Build())
{
    //var column = mailingReportDataTable.GetColumn("Personal Information");
    //mailingReportDataTable.GetColumn(column.SpannedLastColumnNumber).ColumnStyle.RightBorder = XLBorderStyleValues.Double;
    // Specify the file path where you want to save the Excel file
    string filePath = "DataTableToExcel-3-Mailing-List-Report.xlsx";
    d.SaveAsFile(filePath);
}