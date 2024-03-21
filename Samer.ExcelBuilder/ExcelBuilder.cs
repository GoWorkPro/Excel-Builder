using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using static GoWorkPro.ExcelBuilder.ExcelTable;

namespace GoWorkPro.ExcelBuilder
{
    /// <summary>
    /// <author>Samer Shahbaz</author>
    /// <createDate>06/11/2023</createDate>
    /// <changes-lies>Sprint 35</changes-lies>
    /// <description>This utlitiy is developed on the top of ClosedXML.Excel Package. It is an excel wrapper used to develope "Excel Sheets" easily rather than having deep understanding of ClosedXML.Excel.</description>
    /// <contactus-email>info@goworkpro.com</contactus-email>
    /// </summary>
    public class ExcelBuilder : IDisposable, IExcelExtractor, IExcelBuilder
    {
        private bool _isBuild = false;

        public XLWorkbook Workbook { get; }

        public Worksheet[] Worksheets { get; }

        private ExcelBuilder(params Worksheet[] worksheets)
        {
            Workbook = new XLWorkbook();
            Worksheets = worksheets;
            _initializeComponenents();
        }

        private void _initializeComponenents()
        {
            Worksheet[] worksheets = Worksheets;
            foreach (Worksheet worksheet in worksheets)
            {
                Workbook.AddWorksheet(worksheet.Name);
            }
        }

        public static IExcelBuilder Datasets(params ExcelTable[] tables)
        {
            return new ExcelBuilder(new Worksheet("data", tables));
        }

        public static IExcelBuilder Datasets(string workSheetName, params ExcelTable[] tables)
        {
            return new ExcelBuilder(new Worksheet(workSheetName, tables));
        }

        public static IExcelBuilder Datasets(params Worksheet[] worksheets)
        {
            return new ExcelBuilder(worksheets);
        }

        public IExcelExtractor Build()
        {
            if (!_isBuild) _isBuild = true;
            foreach (var builderworksheet in Worksheets)
            {
                var worksheet = Workbook.Worksheet(builderworksheet.Name);

                var nextTableFromRow = 0;
                var currentTableIndex = 0;

                //wast 400
                foreach (var table in builderworksheet.Tables)
                {
                    table.IsBuild = true;
                    //added rows + column's row
                    var totalRows = table.Rows.Count + 1;

                    //updating column
                    for (int columnNumber = 1; columnNumber <= table.MaxCellsCount; columnNumber++)
                    {
                        var columnName = "column" + columnNumber;

                        if (table.Rows.Count > 0 && table.Rows[0].Cells.Count >= columnNumber && !string.IsNullOrWhiteSpace(table.Rows[0].Cells[columnNumber - 1].Value))
                            columnName = table.Rows[0].Cells[columnNumber - 1].Value;

                        table.Columns.AddColumn(columnName, "column" + columnNumber, new ColumnStyle());
                    }


                    for (var i = 0; i < table.EmptyRowsBeforePresentation; i++)
                    {
                        table.Rows.Insert(0, new ExcelRow() { RowStyle = new RowStyle(false) });
                    }

                    // Iterate through the DataTable and add data to the worksheet
                    var rowNumber = nextTableFromRow + 1;
                    table.StartingRowNumber = rowNumber;

                    int cellToAdd = 0;
                    if (table.AlignTableEnd)
                    {
                        if (currentTableIndex > 0)
                        {
                            var prevTable = builderworksheet.Tables.ElementAt(currentTableIndex - 1);
                            int prevMaxCount = Math.Max(prevTable.Rows.OrderByDescending(x => x.Cells.Count).Select(x => x.Cells.Count).FirstOrDefault(), prevTable.Columns.Count);
                            cellToAdd = Math.Max(prevMaxCount - table.MaxCellsCount, 0);
                        }
                    }

                    table.StartColumnNumber = 1 + cellToAdd;

                    var currentTableRowNumber = 0;
                    foreach (ExcelRow row in table.Rows)
                    {
                        var cellNumber = 1 + cellToAdd;
                        var rowStyle = row.RowStyle;

                        //updating row number
                        row.RowNumber = rowNumber;

                        var rowToBe = worksheet.Row(rowNumber);
                        if (rowStyle != null)
                        {
                            if (rowStyle.Height > 0)
                                worksheet.Row(rowNumber).Height = rowStyle.Height;
                            rowStyle.UpdateStyleTo(worksheet.Range(rowNumber, cellNumber, rowNumber, cellToAdd + row.Cells.Count).Style);
                        }

                        if (row.SpanRowToMaxColumns == true)
                        {
                            var maxCellsCount = table.MaxCellsCount;
                            if (table.LinkedTable != null)
                            {
                                maxCellsCount = Math.Max(table.MaxCellsCount, table.LinkedTable.MaxCellsCount);
                            }
                            var range = worksheet.Range(rowNumber, 1, rowNumber, maxCellsCount).Merge();
                            rowStyle.UpdateStyleTo(range.Style);
                        }

                        var cellIndex = 0;
                        foreach (ExcelCell cell in row.Cells)
                        {
                            table.Columns[cellIndex].ColumnNumber = cellNumber;
                            if (currentTableRowNumber <= 0)
                            {
                                table.Columns[cellIndex].SpannedLastColumnNumber = cellNumber + (cell.CellStyle.Colspan > 0 ? cell.CellStyle.Colspan - 1 : 0);
                            }
                            else if (table.Columns[cellIndex].SpannedLastColumnNumber == 0)
                            {
                                table.Columns[cellIndex].SpannedLastColumnNumber = cellNumber;
                            }

                            //updating cell number
                            cell.CellNumber = cellNumber;
                            if (cell.CellStyle.Colspan > 0 || cell.CellStyle.Rowspan > 0)
                            {
                                var range = worksheet.Range(rowNumber, cellNumber, rowNumber + (cell.CellStyle.Rowspan > 0 ? cell.CellStyle.Rowspan - 1 : 0), cellNumber + (cell.CellStyle.Colspan > 0 ? cell.CellStyle.Colspan - 1 : 0)).Merge();
                                rowStyle.UpdateStyleTo(range.Style);
                            }

                            var cellToBe = rowToBe.Cell(cellNumber);

                            _excelCellValue(cell, cellToBe, table);

                            //overriding some styles, comming from row styles
                            if (cell.CellStyle.AllowCellsToBeBordered.HasValue)
                                cell.CellStyle.AllowCellsToBeBordered = rowStyle.AllowCellsToBeBordered;
                            //end

                            cell.CellStyle.UpdateStyleTo(cellToBe.Style);

                            if (builderworksheet.ColumnsWidth.Count - 1 < cellNumber - 1)
                                builderworksheet.ColumnsWidth.Add(0);

                            if (cell.Value.Length > builderworksheet.ColumnsWidth[cellNumber - 1] && !row.SpanRowToMaxColumns && cell.CellStyle.Colspan <= 1)
                            {
                                builderworksheet.ColumnsWidth[cellNumber - 1] = cell.Value.Length;
                            }

                            table.LastColumnNumber = cellNumber;
                            cellIndex++;
                            cellNumber++;
                        }


                        if (row.IsFrozen)
                        {
                            worksheet.SheetView.FreezeRows(rowNumber);
                        }

                        table.LastRowNumber = rowNumber;
                        rowNumber++;
                        currentTableRowNumber++;
                    }

                    currentTableIndex++;
                    nextTableFromRow = rowNumber - 1;
                }

            }
            return this;
        }

        private void _excelCellValue(ExcelTable.ExcelCell sourceCell, IXLCell targetCell, ExcelTable table)
        {
            if (sourceCell.RichText.ContainValue())
            {
                IXLRichText iXLRichText = targetCell.CreateRichText();
                {
                    foreach (ExcelTable.RichTextValue value in sourceCell.RichText.GetValues())
                    {
                        if (value.Value.StartsWith("\n"))
                        {
                            iXLRichText.AddNewLine();
                        }

                        IXLRichString xLRichString = iXLRichText.AddText(value.Value.Replace("\n", ""));
                        value.UpdateStyleTo(ref xLRichString);
                    }

                    return;
                }
            }

            if (table.AllowHtmlToBeParsed)
            {
                _parseHtmlInValue(targetCell, sourceCell.Value);
            }
            else
            {
                targetCell.Value = sourceCell.Value;
            }
        }

        private void _parseHtmlInValue(IXLCell targetCell, string sourceValue)
        {
            string sourceValue2 = sourceValue;
            string[] array = new string[6] { "<b>", "<u>", "<i>", "<br />", "<br/>", "<br>" };
            if (array.Any((string character) => sourceValue2.Contains(character)))
            {
                IXLRichText iXLRichText = targetCell.CreateRichText();
                string[] array2 = array;
                foreach (string text in array2)
                {
                    if (!sourceValue2.Contains(text))
                    {
                        continue;
                    }

                    string[] array3 = sourceValue2.Split(new string[1] { text }, StringSplitOptions.None);
                    int num = 1;
                    string[] array4 = array3;
                    foreach (string text2 in array4)
                    {
                        if (!string.IsNullOrEmpty(text2))
                        {
                            if (text == "<br />" || text == "<br>" || text == "<br/>")
                            {
                                iXLRichText.AddText(text2);
                                if (num != array3.Length)
                                {
                                    iXLRichText.AddNewLine();
                                }
                            }
                            else
                            {
                                switch (text)
                                {
                                    case "<b>":
                                        iXLRichText.AddText(text2).Bold = true;
                                        break;
                                    case "<u>":
                                        iXLRichText.AddText(text2).Underline = XLFontUnderlineValues.Single;
                                        break;
                                    case "<i>":
                                        iXLRichText.AddText(text2).Italic = true;
                                        break;
                                }
                            }
                        }

                        num++;
                    }
                }
            }
            else
            {
                targetCell.Value = sourceValue2;
            }
        }

        public Worksheet GetWorksheet(string name)
        {
            string name2 = name;
            return Worksheets.FirstOrDefault((Worksheet x) => x.Name == name2);
        }

        public Worksheet GetFirstWorksheet()
        {
            return Worksheets.First();
        }

        private void _updateSets()
        {
            Worksheet[] worksheets = Worksheets;
            foreach (Worksheet worksheet in worksheets)
            {
                IXLWorksheet iXLWorksheet = Workbook.Worksheet(worksheet.Name);
                worksheet.UpdateColumnWidths(iXLWorksheet);
                iXLWorksheet.ShowGridLines = worksheet.ShowGridLines;
                foreach (ExcelTable table in worksheet.Tables)
                {
                    foreach (ExcelTable.ExcelColumn column in table.Columns)
                    {
                        IXLRange iXLRange = iXLWorksheet.Range(table.StartingRowNumber + table.EmptyRowsBeforePresentation, column.ColumnNumber, table.LastRowNumber, column.SpannedLastColumnNumber);
                        column.ColumnStyle.UpdateStyleTo(iXLRange.Style);
                        if (column.ColumnStyle.Width > 0)
                        {
                            iXLWorksheet.Column(column.ColumnNumber).Width = column.ColumnStyle.Width;
                        }

                        if (column.IsFrozen)
                        {
                            iXLWorksheet.SheetView.FreezeColumns(column.ColumnNumber);
                        }
                    }
                }
            }
        }

        public IDisposable SaveAsFile(string fileName)
        {
            _updateSets();
            Workbook.SaveAs(fileName);
            return this;
        }

        public MemoryStream GetStream()
        {
            _updateSets();
            MemoryStream memoryStream = new MemoryStream();
            Workbook.SaveAs(memoryStream);
            memoryStream.Seek(0L, SeekOrigin.Begin);
            return memoryStream;
        }

        public void Dispose()
        {
            if (Workbook != null)
            {
                Workbook.Dispose();
            }
        }

        internal static void _checkAndThrowException(bool isBuild, string propertyName)
        {
            if (!isBuild)
            {
                throw new Exception("ExcelTable's " + propertyName + " can not be called before Build() Method.");
            }
        }
    }


    public class Worksheet
    {
        public Worksheet(string name, params ExcelTable[] tables)
        {
            Name = name;
            Tables = tables;
        }

        public bool ShowGridLines { get; set; } = false;
        public string Name { get; }
        public IEnumerable<ExcelTable> Tables { get; }
        public List<double> ColumnsWidth
        {
            get;
            set;
        } = new List<double>();
        internal void UpdateColumnWidths(IXLWorksheet xLWorksheet)
        {
            for (var i = 0; i < ColumnsWidth.Count; i++)
            {
                var column = xLWorksheet.Column(i + 1);
                if (ColumnsWidth[i] > 0)
                    column.Width = (ColumnsWidth[i] + 1) > 50 ? 50 : (ColumnsWidth[i] + 1);
            }
        }

    }
}
