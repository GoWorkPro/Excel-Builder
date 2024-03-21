using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace GoWorkPro.ExcelBuilder
{
    public class ExcelTable
    {
        public class ExcelRow
        {
            public List<ExcelCell> Cells { get; set; }

            public RowStyle RowStyle { get; set; }

            public bool SpanRowToMaxColumns { get; private set; }

            public bool IsFrozen { get; private set; }

            public int RowNumber { get; internal set; }

            public ExcelRow()
            {
                Cells = new List<ExcelCell>();
                RowStyle = new RowStyle();
            }

            public ExcelRow(params string[] cellsValue)
                : this()
            {
                this.AddRange(cellsValue);
            }

            public ExcelRow(RowStyle rowStyle, params string[] cellsValue)
                : this()
            {
                this.AddRange(rowStyle, cellsValue);
            }

            public ExcelRow(RowStyle rowStyle)
                : this()
            {
                RowStyle = rowStyle;
            }

            public void SpanToMaxRowCells()
            {
                SpanRowToMaxColumns = true;
            }

            public void Freeze()
            {
                IsFrozen = true;
            }
        }

        public class ExcelColumn
        {
            public string ColumnName { get; }

            public string ActualName { get; }

            public ColumnStyle ColumnStyle { get; set; }

            public int ColumnNumber { get; internal set; }

            public int SpannedLastColumnNumber { get; internal set; }
            public bool IsFrozen { get; internal set; }

            public void Freeze()
            {
                IsFrozen = true;
            }

            public ExcelColumn(string columnName, string actualColumnName)
            {
                ColumnStyle = new ColumnStyle();
                ColumnName = columnName;
                ActualName = actualColumnName;
            }

            public ExcelColumn(string columnName, string actualColumnName, ColumnStyle columnStyle)
                : this(columnName, actualColumnName)
            {
                ColumnStyle = columnStyle;
            }
        }

        public class ExcelStyle
        {
            public bool? FontBold { get; set; }

            public XLAlignmentHorizontalValues? AlignmentHorizontal { get; set; }

            public XLAlignmentVerticalValues? AlignmentVertical { get; set; } = XLAlignmentVerticalValues.Center;


            public XLColor? BackgroundColor { get; set; }

            public XLColor? FontColor { get; set; }

            public double FontSize { get; set; }

            public XLBorderStyleValues? TopBorder { get; set; }

            public XLBorderStyleValues? BottomBorder { get; set; }

            public XLBorderStyleValues? LeftBorder { get; set; }

            public XLBorderStyleValues? RightBorder { get; set; }

            public XLColor TopBorderColor { get; set; }
            public XLColor LeftBorderColor { get; set; }
            public XLColor BottomBorderColor { get; set; }
            public XLColor RightBorderColor { get; set; }


            public bool AddCellsToSpan { get; set; } = true;


            public bool? AllowCellsToBeBordered { get; set; }

            public ExcelStyle()
            {
            }

            public ExcelStyle(bool? allowCellsToBeBordered = null)
            {
                AllowCellsToBeBordered = allowCellsToBeBordered;
            }

            public virtual void UpdateStyleTo(IXLStyle xLStyle)
            {
                if (AlignmentHorizontal.HasValue)
                {
                    xLStyle.Alignment.Horizontal = AlignmentHorizontal.Value;
                }

                if (AlignmentVertical.HasValue)
                {
                    xLStyle.Alignment.Vertical = AlignmentVertical.Value;
                }

                if (FontSize != 0.0)
                {
                    xLStyle.Font.FontSize = FontSize;
                }

                if (FontBold.HasValue)
                {
                    xLStyle.Font.Bold = FontBold.Value;
                }

                if (FontColor != null)
                {
                    xLStyle.Font.FontColor = FontColor;
                }

                if (BackgroundColor != null)
                {
                    xLStyle.Fill.BackgroundColor = BackgroundColor;
                }

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

        public class CellStyle : ExcelStyle, ICloneable
        {
            public virtual int Colspan { get; set; }

            public virtual int Rowspan { get; set; }

            public bool IsWrapped { get; set; } = true;
            internal bool InMergeCell { get; set; } = false;

            public CellStyle(bool? allowCellsToBeBordered = null)
            {
                base.AllowCellsToBeBordered = allowCellsToBeBordered;
            }


            public override void UpdateStyleTo(IXLStyle xLStyle)
            {
                base.UpdateStyleTo(xLStyle);
                xLStyle.Alignment.WrapText = IsWrapped;
            }

            public object Clone()
            {
                return new CellStyle
                {
                    InMergeCell = InMergeCell,
                    Rowspan = Rowspan,
                    AddCellsToSpan = AddCellsToSpan,
                    AlignmentHorizontal = AlignmentHorizontal,
                    AlignmentVertical = AlignmentVertical,
                    AllowCellsToBeBordered = AllowCellsToBeBordered,
                    BackgroundColor = BackgroundColor,
                    BottomBorder = BottomBorder,
                    BottomBorderColor = BottomBorderColor,
                    Colspan = Colspan,
                    FontBold = FontBold,
                    FontColor = FontColor,
                    FontSize = FontSize,
                    IsWrapped = IsWrapped,
                    LeftBorder = LeftBorder,
                    LeftBorderColor = LeftBorderColor,
                    RightBorder = RightBorder,
                    RightBorderColor = RightBorderColor,
                    TopBorder = TopBorder,
                    TopBorderColor = TopBorderColor
                };
            }
        }

        public class RowStyle : ExcelStyle
        {
            public double Height { get; set; }

            public RowStyle(bool allowCellsToBeBordered = true)
            {
                base.AllowCellsToBeBordered = allowCellsToBeBordered;
            }
        }

        public class ExcelCell
        {
            public CellStyle CellStyle { get; set; }

            public string Value { get; set; } = "";


            public int CellNumber { get; internal set; }

            public RichText RichText { get; set; }

            public ExcelCell()
            {
                CellStyle = new CellStyle();
                RichText = new RichText();
            }

            public ExcelCell(CellStyle cellStyle)
                : this()
            {
                CellStyle = cellStyle;
            }

            public ExcelCell(string value)
                : this()
            {
                Value = value;
            }

            public ExcelCell(RichText richText)
                : this()
            {
                RichText = richText;
            }

            public ExcelCell(RichText richText, CellStyle cellStyle)
                : this()
            {
                RichText = richText;
                CellStyle = cellStyle;
            }

            public ExcelCell(string value, CellStyle cellStyle)
                : this()
            {
                Value = value;
                CellStyle = cellStyle;
            }
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
                {
                    xLRichString.FontColor = FontColor;
                }
            }
        }

        public class RichText
        {
            private List<RichTextValue> _values;

            public RichText()
            {
                _values = new List<RichTextValue>();
            }

            public bool ContainValue()
            {
                return _values.Any();
            }

            public void Add(RichTextValue richTextValue)
            {
                _values.Add(richTextValue);
            }

            public void AddALine(RichTextValue richTextValue)
            {
                richTextValue.Value = "\n" + richTextValue.Value;
                _values.Add(richTextValue);
            }

            public IEnumerable<RichTextValue> GetValues()
            {
                return _values;
            }
        }

        internal bool IsBuild;

        public bool AlignTableEnd { get; set; }

        public int EmptyRowsBeforePresentation { get; set; }

        public List<ExcelColumn> Columns { get; internal set; }

        public List<ExcelRow> Rows { get; set; }

        public int MaxCellsCount => (from x in Rows
                                     orderby x.Cells.Count descending
                                     select x.Cells.Count).FirstOrDefault();

        public int StartColumnNumber { get; internal set; }

        public int LastColumnNumber { get; internal set; }

        public int StartingRowNumber { get; internal set; }

        public int LastRowNumber { get; internal set; }

        public bool AllowHtmlToBeParsed { get; set; } = true;


        public ExcelTable? LinkedTable { get; set; }

        public ExcelTable()
        {
            Columns = new List<ExcelColumn>();
            Rows = new List<ExcelRow>();
        }

        public ExcelColumn? GetColumn(string columnName)
        {
            string columnName2 = columnName;
            _checkAndThrowException("GetColumn");
            return Columns.FirstOrDefault((ExcelColumn x) => x.ColumnName == columnName2 || x.ActualName == columnName2);
        }

        public ExcelColumn GetColumn(int columnNumber)
        {
            _checkAndThrowException("GetColumn");
            return Columns.FirstOrDefault((ExcelColumn x) => x.ColumnNumber == columnNumber);
        }

        private void _checkAndThrowException(string propertyName)
        {
            ExcelBuilder._checkAndThrowException(IsBuild, propertyName);
        }

        public void Link(ExcelTable linkTable)
        {
            LinkedTable = linkTable;
        }
    }
}
