using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static GoWorkPro.ExcelBuilder.ExcelTable;

namespace GoWorkPro.ExcelBuilder
{
    #region Excel Utility Extention
    public static class ExcelCollectionEx
    {
        public static void AddCell(this ICollection<ExcelTable.ExcelCell> cells, string value)
        {
            cells.Add(new ExcelTable.ExcelCell
            {
                Value = value
            });
        }

        public static void AddCell(this ICollection<ExcelTable.ExcelCell> cells, string value, ExcelTable.CellStyle cellStyle)
        {
            cells.Add(new ExcelTable.ExcelCell
            {
                Value = value,
                CellStyle = cellStyle
            });
            if (cellStyle.AddCellsToSpan)
            {
                for (int i = 0; i < cellStyle.Colspan - 1; i++)
                {
                    cells.Add(new ExcelTable.ExcelCell());
                }
            }
        }

        public static void AddCell(this ExcelTable.ExcelRow row, string value, ExcelTable.CellStyle cellStyle)
        {
            row.Cells.Add(new ExcelTable.ExcelCell
            {
                Value = value,
                CellStyle = cellStyle
            });
            if (cellStyle.AddCellsToSpan)
            {
                for (int i = 0; i < cellStyle.Colspan - 1; i++)
                {
                    row.AddCell();
                }
            }
        }

        public static void AddCell(this ExcelTable.ExcelRow row, ExcelTable.ExcelCell cell)
        {
            row.Cells.Add(cell);
            if (cell.CellStyle.AddCellsToSpan)
            {
                for (int i = 0; i < cell.CellStyle.Colspan - 1; i++)
                {
                    row.AddCell();
                }
            }
        }

        public static string[] Values(this ExcelTable.ExcelRow excelRow)
        {
            return excelRow.Cells.Select((ExcelTable.ExcelCell x) => x.Value).ToArray();
        }

        internal static void AddColumn(this ICollection<ExcelTable.ExcelColumn> columns, string columnName, string actualColumnName, ExcelTable.ColumnStyle columnStyle)
        {
            columns.Add(new ExcelTable.ExcelColumn(columnName, actualColumnName, columnStyle));
        }

        public static void AddCell(this ExcelTable.ExcelRow row, string value = "")
        {
            row.AddCell(value, new ExcelTable.CellStyle());
        }

        public static ExcelTable.ExcelRow Add(this ICollection<ExcelTable.ExcelRow> rows, params string[] cellsValue)
        {
            List<ExcelTable.ExcelCell> list = new List<ExcelTable.ExcelCell>();
            foreach (string value in cellsValue)
            {
                list.Add(new ExcelTable.ExcelCell
                {
                    Value = value
                });
            }

            ExcelTable.ExcelRow excelRow = new ExcelTable.ExcelRow
            {
                Cells = list
            };
            rows.Add(excelRow);
            return excelRow;
        }

        public static ExcelTable.ExcelRow Add(this ICollection<ExcelTable.ExcelRow> rows)
        {
            List<ExcelTable.ExcelCell> cells = new List<ExcelTable.ExcelCell>();
            ExcelTable.ExcelRow excelRow = new ExcelTable.ExcelRow
            {
                Cells = cells,
                RowStyle = new ExcelTable.RowStyle(allowCellsToBeBordered: false)
            };
            rows.Add(excelRow);
            return excelRow;
        }

        public static ExcelTable.ExcelRow Add(this ICollection<ExcelTable.ExcelRow> rows, ExcelTable.RowStyle rowStyle, params string[] cellsValue)
        {
            ExcelTable.ExcelRow excelRow = new ExcelTable.ExcelRow
            {
                RowStyle = rowStyle
            };
            foreach (string value in cellsValue)
            {
                excelRow.AddCell(value);
            }

            rows.Add(excelRow);
            return excelRow;
        }

        public static ExcelTable.ExcelRow Add(this ICollection<ExcelTable.ExcelRow> rows, ExcelTable.CellStyle cellStyle, params string[] cellsValue)
        {
            ExcelTable.ExcelRow excelRow = new ExcelTable.ExcelRow();
            foreach (string value in cellsValue)
            {
                excelRow.AddCell(value, cellStyle);
            }

            rows.Add(excelRow);
            return excelRow;
        }

        public static ExcelTable.ExcelRow Add(this ICollection<ExcelTable.ExcelRow> rows, ExcelTable.RowStyle rowStyle, ExcelTable.CellStyle cellStyle, params string[] cellsValue)
        {
            ExcelTable.ExcelRow excelRow = rows.Add(cellStyle, cellsValue);
            excelRow.RowStyle = rowStyle;
            return excelRow;
        }

        public static ExcelTable.ExcelRow Add(this ICollection<ExcelTable.ExcelRow> rows, params ExcelTable.ExcelCell[] excelCells)
        {
            ExcelTable.ExcelRow excelRow = new ExcelTable.ExcelRow();
            foreach (ExcelTable.ExcelCell cell in excelCells)
            {
                excelRow.AddCell(cell);
            }

            rows.Add(excelRow);
            return excelRow;
        }

        public static ExcelTable.ExcelRow Add(this ICollection<ExcelTable.ExcelRow> rows, ExcelTable.RowStyle rowStyle, params ExcelTable.ExcelCell[] excelCells)
        {
            ExcelTable.ExcelRow excelRow = rows.Add(excelCells);
            excelRow.RowStyle = rowStyle;
            return excelRow;
        }

        public static void AddRange(this ICollection<ExcelTable.ExcelCell> cells, ExcelTable.CellStyle cellStyle, params string[] cellsValue)
        {
            foreach (string value in cellsValue)
            {
                cells.Add(new ExcelTable.ExcelCell
                {
                    Value = value,
                    CellStyle = cellStyle
                });
                if (cellStyle.AddCellsToSpan)
                {
                    for (int j = 0; j < cellStyle.Colspan - 1; j++)
                    {
                        cells.Add(new ExcelTable.ExcelCell
                        {
                            CellStyle = cellStyle
                        });
                    }
                }
            }
        }

        public static void AddRange(this ExcelTable.ExcelRow row, params string[] cellsValue)
        {
            List<ExcelTable.ExcelCell> list = new List<ExcelTable.ExcelCell>();
            foreach (string value in cellsValue)
            {
                list.Add(new ExcelTable.ExcelCell
                {
                    Value = value
                });
            }

            row.Cells = list;
        }

        public static void AddRange(this ExcelTable.ExcelRow row, ExcelTable.RowStyle rowStyle, params string[] cellsValue)
        {
            List<ExcelTable.ExcelCell> list = new List<ExcelTable.ExcelCell>();
            foreach (string value in cellsValue)
            {
                list.Add(new ExcelTable.ExcelCell
                {
                    Value = value
                });
            }

            row.RowStyle = rowStyle;
            row.Cells = list;
        }
    }
    #endregion
}
