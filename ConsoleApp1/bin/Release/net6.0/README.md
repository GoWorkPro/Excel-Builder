ExcelBuilder
Author: Samer Shahbaz
Create Date: 06/11/2023

Description
The ExcelBuilder utility is developed on top of the ClosedXML.Excel package, providing a convenient wrapper for creating Excel sheets without requiring an in-depth understanding of ClosedXML.Excel.

Features
Simplified Excel Sheet Development: Easily create Excel sheets without extensive knowledge of ClosedXML.Excel.
Table Creation: Build tables with rows, columns, and customizable styles.
Rich Text Support: Support for rich text formatting within cells.
Cell and Column Styling: Customize cell and column styles, including font, alignment, borders, and more.
Flexibility: Build multiple worksheets with different tables in a single workbook.
Usage
Create ExcelBuilder Instance:
var excelBuilder = ExcelBuilder.Datasets(
    new ExcelTable(
        new ExcelRow("Header1", "Header2"),
        new ExcelRow("Data1", "Data2")
    )
);

Custom Styling:
var customStyle = new ExcelTable.CellStyle
{
    FontBold = true,
    FontColor = XLColor.Blue,
    // ... other style configurations
};

var excelBuilder = ExcelBuilder.Datasets(
    new ExcelTable(
        new ExcelRow(customStyle, "Bold Blue Header", "Another Header"),
        new ExcelRow("Data1", "Data2")
    )
);

Save to File:
using (var file = excelBuilder.SaveAsFile("example.xlsx"))
{
    // Perform additional operations or return the file stream
}

Retrieve Worksheet:
var worksheet = excelBuilder.GetWorksheet("Sheet1");

Stream Output:
var excelStream = excelBuilder.GetStream();
// Use the stream as needed

Important Notes
At least one table is required to build an Excel sheet.
The utility is built on the ClosedXML.Excel package (version: 1.0.0).

See more details on our project site
Feel free to contribute, report issues, or suggest improvements!