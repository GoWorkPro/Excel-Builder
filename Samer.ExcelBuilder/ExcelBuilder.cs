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
            foreach (var builderWorksheet in Worksheets)
            {
                Workbook.AddWorksheet(builderWorksheet.Name);
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

        bool _isBuild = false;

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
                            cellToAdd = prevMaxCount - table.MaxCellsCount;
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

                            _excelCellValue(cell, cellToBe);

                            //overriding some styles, comming from row styles
                            if (cell.CellStyle.AllowCellsToBeBordered == null)
                                cell.CellStyle.AllowCellsToBeBordered = rowStyle.AllowCellsToBeBordered;
                            //end

                            cell.CellStyle.UpdateStyleTo(cellToBe.Style);

                            if (builderworksheet.ColumnsWidth.Count - 1 < cellNumber - 1)
                                builderworksheet.ColumnsWidth.Add(0);

                            if (cell.Value.Length > builderworksheet.ColumnsWidth[cellNumber - 1])
                            {
                                if (row.SpanRowToMaxColumns == false && cell.CellStyle.Colspan <= 1)
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

        private void _excelCellValue(ExcelCell sourceCell, IXLCell targetCell)
        {
            var isRichText = sourceCell.RichText.HashValue();

            if (isRichText)
            {
                var richText = targetCell.CreateRichText();
                foreach (var rich in sourceCell.RichText.GetValues())
                {
                    if (rich.Value.StartsWith("\n"))
                        richText.AddNewLine();

                    IXLRichString xLRichString = richText.AddText(rich.Value.Replace("\n", ""));
                    rich.UpdateStyleTo(ref xLRichString);
                }
            }
            else
            {
                _parseHtmlInValue(targetCell, sourceCell.Value);
            }
        }

        void _parseHtmlInValue(IXLCell targetCell, string sourceValue)
        {
            var supportedHtmlCharacters = new string[] { "<b>", "<u>", "<i>", "<br />", "<br/>", "<br>" };


            if (supportedHtmlCharacters.Any(character => sourceValue.Contains(character)))
            {
                var richText = targetCell.CreateRichText();
                foreach (var supportedHtmlCharacter in supportedHtmlCharacters)
                {
                    if (sourceValue.Contains(supportedHtmlCharacter))
                    {
                        // Split and process based on supported HTML characters
                        var parts = sourceValue.Split(new[] { supportedHtmlCharacter }, StringSplitOptions.None);

                        var partNumber = 1;
                        foreach (var part in parts)
                        {
                            if (!string.IsNullOrEmpty(part))
                            {
                                if (supportedHtmlCharacter == "<br />" || supportedHtmlCharacter == "<br>" || supportedHtmlCharacter == "<br/>")
                                {
                                    richText.AddText(part);
                                    if (partNumber != parts.Length)
                                        richText.AddNewLine();
                                }
                                else
                                {
                                    // Handle other HTML tags
                                    if (supportedHtmlCharacter == "<b>")
                                    {
                                        richText.AddText(part).Bold = true;
                                    }
                                    else if (supportedHtmlCharacter == "<u>")
                                    {
                                        richText.AddText(part).Underline = XLFontUnderlineValues.Single;
                                    }
                                    else if (supportedHtmlCharacter == "<i>")
                                    {
                                        richText.AddText(part).Italic = true;
                                    }
                                }
                            }
                            partNumber++;
                        }
                    }
                }
            }
            else
            {
                // If none of the supported HTML characters found, add the text as it is
                targetCell.Value = sourceValue;
            }
        }

        public Worksheet GetWorksheet(string name)
        {
            return Worksheets.FirstOrDefault(x => x.Name == name);
        }

        public Worksheet GetFirstWorksheet()
        {
            return Worksheets.First();
        }

        void _updateSets()
        {
            foreach (var worksheetBuilder in Worksheets)
            {
                var worksheet = Workbook.Worksheet(worksheetBuilder.Name);
                worksheetBuilder.UpdateColumnWidths(worksheet);
                worksheet.ShowGridLines = worksheetBuilder.ShowGridLines;

                foreach (var table in worksheetBuilder.Tables)
                {
                    foreach (var column in table.Columns)
                    {
                        var columnRange = worksheet.Range(table.StartingRowNumber + table.EmptyRowsBeforePresentation, column.ColumnNumber, table.LastRowNumber, column.SpannedLastColumnNumber);
                        column.ColumnStyle.UpdateStyleTo(columnRange.Style);

                        if (column.ColumnStyle.Width > 0)
                            worksheet.Column(column.ColumnNumber).Width = column.ColumnStyle.Width;
                    }
                }
            }
        }

        public IDisposable SaveAsFile(string fileName)
        {
            _updateSets();
            this.Workbook.SaveAs(fileName);
            return this;
        }

        public MemoryStream GetStream()
        {
            _updateSets();
            var excelStream = new MemoryStream();
            Workbook.SaveAs(excelStream);
            excelStream.Seek(0, SeekOrigin.Begin);
            return excelStream;
        }

        public void Dispose()
        {
            if (Workbook != null)
                Workbook.Dispose();
        }

        internal static void _checkAndThrowException(bool isBuild, string propertyName)
        {
            if (!isBuild)
                throw new Exception($"ExcelTable's {propertyName} can not be called before Build() Method.");
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

    #region Excel objects to build
    public class ExcelTable
    {
        public ExcelTable()
        {
            Columns = new List<ExcelColumn>();
            Rows = new List<ExcelRow>();
        }
        public bool AlignTableEnd { get; set; }
        public int EmptyRowsBeforePresentation { get; set; }
        public List<ExcelColumn> Columns { get; internal set; }
        public List<ExcelRow> Rows { get; set; }
        public int MaxCellsCount { get => this.Rows.OrderByDescending(x => x.Cells.Count).Select(x => x.Cells.Count).FirstOrDefault(); }
        public int StartColumnNumber { get; internal set; }
        public int LastColumnNumber { get; internal set; }
        public int StartingRowNumber { get; internal set; }
        public int LastRowNumber { get; internal set; }
        public ExcelTable? LinkedTable { get; set; }
        internal bool IsBuild;
        public ExcelColumn? GetColumn(string columnName)
        {
            _checkAndThrowException(nameof(GetColumn));
            return this.Columns.FirstOrDefault(x => x.ColumnName == columnName || x.ActualName == columnName);
        }


        public ExcelColumn GetColumn(int columnNumber)
        {
            _checkAndThrowException(nameof(GetColumn));
            return this.Columns.FirstOrDefault(x => x.ColumnNumber == columnNumber);
        }



        private void _checkAndThrowException(string propertyName)
        {
            ExcelBuilder._checkAndThrowException(this.IsBuild, propertyName);
        }

        public void Link(ExcelTable linkTable)
        {
            LinkedTable = linkTable;
        }

        public class ExcelRow
        {
            public ExcelRow()
            {
                Cells = new List<ExcelCell>();
                RowStyle = new RowStyle();
            }

            public ExcelRow(params string[] cellsValue) : this()
            {
                this.AddRange(cellsValue);
            }

            public ExcelRow(RowStyle rowStyle, params string[] cellsValue) : this()
            {
                this.AddRange(rowStyle, cellsValue);
            }

            public ExcelRow(RowStyle rowStyle) : this()
            {
                this.RowStyle = rowStyle;
            }

            public void SpanToMaxRowCells()
            {
                this.SpanRowToMaxColumns = true;
            }

            public List<ExcelCell> Cells { get; set; }
            public RowStyle RowStyle { get; set; }
            public bool SpanRowToMaxColumns { get; private set; }
            public bool IsFrozen { get; private set; }
            public int RowNumber { get; internal set; }
            public void Freeze()
            {
                this.IsFrozen = true;
            }
        }

        public class ExcelColumn
        {
            public ExcelColumn(string columnName, string actualColumnName)
            {
                ColumnStyle = new ColumnStyle();
                ColumnName = columnName;
                ActualName = actualColumnName;
            }

            public ExcelColumn(string columnName, string actualColumnName, ColumnStyle columnStyle) : this(columnName, actualColumnName)
            {
                ColumnStyle = columnStyle;
            }

            public string ColumnName { get; }
            public string ActualName { get; }

            public ColumnStyle ColumnStyle { get; set; }

            public int ColumnNumber { get; internal set; }
            public int SpannedLastColumnNumber { get; internal set; }
        }

        public class ExcelStyle
        {
            public ExcelStyle() { }
            public ExcelStyle(bool? allowCellsToBeBordered = null)
            {
                this.AllowCellsToBeBordered = allowCellsToBeBordered;
            }
            public bool? FontBold { get; set; }
            public XLAlignmentHorizontalValues? AlignmentHorizontal { get; set; }
            public XLAlignmentVerticalValues? AlignmentVertical { get; set; } = XLAlignmentVerticalValues.Center;
            public XLColor? BackroundColor { get; set; }
            public XLColor? FontColor { get; set; }
            public double FontSize { get; set; }
            public XLBorderStyleValues? TopBorder { get; set; }
            public XLBorderStyleValues? BottomBorder { get; set; }
            public XLBorderStyleValues? LeftBorder { get; set; }
            public XLBorderStyleValues? RightBorder { get; set; }
            public bool AddCellsToSpan { get; set; } = true;
            public bool? AllowCellsToBeBordered { get; set; }
            public XLColor TopBorderColor { get; set; }
            public XLColor LeftBorderColor { get; set; }
            public XLColor BottomBorderColor { get; set; }
            public XLColor RightBorderColor { get; set; }   
            public virtual void UpdateStyleTo(IXLStyle xLStyle)
            {
                if (AlignmentHorizontal != null)
                    xLStyle.Alignment.Horizontal = (XLAlignmentHorizontalValues)AlignmentHorizontal;
                if (AlignmentVertical != null)
                    xLStyle.Alignment.Vertical = (XLAlignmentVerticalValues)AlignmentVertical;
                if (FontSize != 0)
                    xLStyle.Font.FontSize = FontSize;
                if (FontBold != null)
                    xLStyle.Font.Bold = (bool)FontBold;
                if (FontColor != null)
                    xLStyle.Font.FontColor = FontColor;
                if (BackroundColor != null)
                    xLStyle.Fill.BackgroundColor = BackroundColor;

                if (AllowCellsToBeBordered == true)
                {
                    xLStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    xLStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    xLStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    xLStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    xLStyle.Border.TopBorderColor = XLColor.FromHtml("#D9D9D9");
                    xLStyle.Border.RightBorderColor = XLColor.FromHtml("#D9D9D9");
                    xLStyle.Border.BottomBorderColor = XLColor.FromHtml("#D9D9D9");
                    xLStyle.Border.LeftBorderColor = XLColor.FromHtml("#D9D9D9");
                }

                if (TopBorder != null)
                    xLStyle.Border.TopBorder = (XLBorderStyleValues)TopBorder;

                if (BottomBorder != null)
                    xLStyle.Border.BottomBorder = (XLBorderStyleValues)BottomBorder;

                if (LeftBorder != null)
                    xLStyle.Border.LeftBorder = (XLBorderStyleValues)LeftBorder;

                if (RightBorder != null)
                    xLStyle.Border.RightBorder = (XLBorderStyleValues)RightBorder;


                if (TopBorderColor != null)
                    xLStyle.Border.TopBorderColor = TopBorderColor;
                if (RightBorderColor != null)
                    xLStyle.Border.RightBorderColor = RightBorderColor;
                if (BottomBorderColor != null)
                    xLStyle.Border.BottomBorderColor = BottomBorderColor;
                if (LeftBorderColor != null)
                    xLStyle.Border.LeftBorderColor = LeftBorderColor;


            }
        }

        public class ColumnStyle : ExcelStyle
        {
            public int Width { get; set; }
            public override void UpdateStyleTo(IXLStyle xLStyle)
            {
                base.UpdateStyleTo(xLStyle);
            }
        }

        public class CellStyle : ExcelStyle
        {
            public CellStyle(bool? allowCellsToBeBordered = null)
            {
                AllowCellsToBeBordered = allowCellsToBeBordered;
            }
            public virtual int Colspan { get; set; }
            public virtual int Rowspan { get; set; }
            public bool IsWrapped { get; set; } = true;

            public override void UpdateStyleTo(IXLStyle xLStyle)
            {
                base.UpdateStyleTo(xLStyle);
                xLStyle.Alignment.WrapText = IsWrapped;
            }
        }

        public class RowStyle : ExcelStyle
        {
            public RowStyle(bool allowCellsToBeBordered = true)
            {
                this.AllowCellsToBeBordered = allowCellsToBeBordered;
            }
            public double Height { get; set; }
        }

        public class ExcelCell
        {
            public ExcelCell()
            {
                CellStyle = new CellStyle();
                RichText = new RichText();
            }

            public ExcelCell(CellStyle cellStyle) : this()
            {
                CellStyle = cellStyle;
            }

            public ExcelCell(string value) : this()
            {
                this.Value = value;
            }

            public ExcelCell(RichText richText) : this()
            {
                this.RichText = richText;
            }

            public ExcelCell(RichText richText, CellStyle cellStyle) : this()
            {
                this.RichText = richText;
                CellStyle = cellStyle;
            }

            public ExcelCell(string value, CellStyle cellStyle) : this()
            {
                this.Value = value;
                CellStyle = cellStyle;
            }

            public CellStyle CellStyle { get; set; }
            public string Value { get; set; } = "";
            public int CellNumber { get; internal set; }
            public RichText RichText { get; set; }
        }

        public class RichTextValue
        {
            public bool Bold { get; set; }
            public XLColor? FontColor { get; set; }
            public bool Italic { get; set; }
            public bool StrikeThrough { get; set; }
            public string Value { get; set; } = "";
            public void UpdateStyleTo(ref IXLRichString xLRichString)
            {
                xLRichString.Bold = Bold;
                xLRichString.Italic = Italic;
                xLRichString.Strikethrough = StrikeThrough;

                if (FontColor != null)
                    xLRichString.FontColor = FontColor;
            }
        }
        public class RichText
        {
            public RichText()
            {
                this._values = new List<RichTextValue>();
            }
            List<RichTextValue> _values;
            public bool HashValue()
            {
                return _values.Any();
            }

            public void Add(RichTextValue richTextValue)
            {
                this._values.Add(richTextValue);
            }

            public void AddALine(RichTextValue richTextValue)
            {
                richTextValue.Value = "\n" + richTextValue.Value;
                this._values.Add(richTextValue);
            }

            public IEnumerable<RichTextValue> GetValues() => _values;
        }

    }
    #endregion

    #region Excel Utility Extention
    public static class ExcelCollectionEx
    {
        public static void AddCell(this ICollection<ExcelCell> cells, string value)
        {
            cells.Add(new ExcelCell
            {
                Value = value,
            });
        }

        public static void AddCell(this ICollection<ExcelCell> cells, string value, ExcelTable.CellStyle cellStyle)
        {
            cells.Add(new ExcelCell
            {
                Value = value,
                CellStyle = cellStyle,
            });

            if (cellStyle.AddCellsToSpan)
                for (int i = 0; i < cellStyle.Colspan - 1; i++)
                {
                    cells.Add(new ExcelCell());
                }
        }

        public static void AddCell(this ExcelRow row, string value, ExcelTable.CellStyle cellStyle)
        {
            row.Cells.Add(new ExcelCell
            {
                Value = value,
                CellStyle = cellStyle,
            });

            if (cellStyle.AddCellsToSpan)
                for (int i = 0; i < cellStyle.Colspan - 1; i++)
                {
                    row.AddCell();
                }
        }


        public static void AddCell(this ExcelRow row, ExcelCell cell)
        {
            row.Cells.Add(cell);

            if (cell.CellStyle.AddCellsToSpan)
                for (int i = 0; i < cell.CellStyle.Colspan - 1; i++)
                {
                    row.AddCell();
                }
        }

        public static string[] Values(this ExcelRow excelRow) => excelRow.Cells.Select(x => x.Value).ToArray();

        internal static void AddColumn(this ICollection<ExcelColumn> columns, string columnName, string actualColumnName, ExcelTable.ColumnStyle columnStyle)
        {
            columns.Add(new ExcelColumn(columnName, actualColumnName, columnStyle));
        }

        public static void AddCell(this ExcelRow row, string value = "")
        {
            AddCell(row, value, new ExcelTable.CellStyle());
        }

        public static ExcelRow Add(this ICollection<ExcelTable.ExcelRow> rows, params string[] cellsValue)
        {
            var rowCells = new List<ExcelTable.ExcelCell>();
            foreach (var value in cellsValue)
            {
                rowCells.Add(new ExcelCell
                {
                    Value = value
                });
            }
            var row = new ExcelRow
            {
                Cells = rowCells
            };

            rows.Add(row);
            return row;
        }

        //for empty
        public static ExcelRow Add(this ICollection<ExcelTable.ExcelRow> rows)
        {
            var rowCells = new List<ExcelTable.ExcelCell>();
            var row = new ExcelRow
            {
                Cells = rowCells,
                RowStyle = new RowStyle(false)
            };

            rows.Add(row);
            return row;
        }



        public static ExcelRow Add(this ICollection<ExcelTable.ExcelRow> rows, RowStyle rowStyle, params string[] cellsValue)
        {
            var row = new ExcelRow
            {
                RowStyle = rowStyle
            };
            foreach (var value in cellsValue)
            {
                AddCell(row, value);
            }

            rows.Add(row);
            return row;
        }

        public static ExcelRow Add(this ICollection<ExcelTable.ExcelRow> rows, ExcelTable.CellStyle cellStyle, params string[] cellsValue)
        {
            var row = new ExcelRow();
            foreach (var value in cellsValue)
            {
                AddCell(row, value, cellStyle);
            }
            rows.Add(row);
            return row;
        }

        public static ExcelRow Add(this ICollection<ExcelTable.ExcelRow> rows, ExcelTable.RowStyle rowStyle, ExcelTable.CellStyle cellStyle, params string[] cellsValue)
        {
            var row = Add(rows, cellStyle, cellsValue);
            row.RowStyle = rowStyle;
            return row;
        }

        public static ExcelRow Add(this ICollection<ExcelRow> rows, params ExcelCell[] excelCells)
        {
            var row = new ExcelRow();
            foreach (var cell in excelCells)
            {
                AddCell(row, cell);
            }

            rows.Add(row);
            return row;
        }

        public static ExcelRow Add(this ICollection<ExcelRow> rows, ExcelTable.RowStyle rowStyle, params ExcelCell[] excelCells)
        {
            var row = Add(rows, excelCells);
            row.RowStyle = rowStyle;
            return row;
        }

        public static void AddRange(this ICollection<ExcelTable.ExcelCell> cells, ExcelTable.CellStyle cellStyle, params string[] cellsValue)
        {
            foreach (var value in cellsValue)
            {
                cells.Add(new ExcelCell
                {
                    Value = value,
                    CellStyle = cellStyle
                });

                if (cellStyle.AddCellsToSpan)
                    for (int i = 0; i < cellStyle.Colspan - 1; i++)
                    {
                        cells.Add(new ExcelCell() { CellStyle = cellStyle });
                    }
            }
        }

        public static void AddRange(this ExcelTable.ExcelRow row, params string[] cellsValue)
        {
            var rowCells = new List<ExcelCell>();
            foreach (var value in cellsValue)
            {
                rowCells.Add(new ExcelCell
                {
                    Value = value
                });
            }
            row.Cells = rowCells;
        }
        public static void AddRange(this ExcelTable.ExcelRow row, RowStyle rowStyle, params string[] cellsValue)
        {
            var rowCells = new List<ExcelCell>();
            foreach (var value in cellsValue)
            {
                rowCells.Add(new ExcelCell
                {
                    Value = value
                });
            }
            row.RowStyle = rowStyle;
            row.Cells = rowCells;
        }

    }
    #endregion
}
