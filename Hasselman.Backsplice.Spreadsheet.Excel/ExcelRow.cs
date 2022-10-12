// SPDX-FileCopyrightText: 2022 James D. Hasselman <james.d.hasselman@gmail.com>
// SPDX-License-Identifier: GPL-3.0-or-later

using DocumentFormat.OpenXml.Spreadsheet;

#if DEBUG
using System.Runtime.CompilerServices;
[assembly: InternalsVisibleToAttribute("UnitTests")]
#endif

namespace Hasselman.Backsplice.Spreadsheet.Excel
{
    public class ExcelRow : IRow
    {
        internal Row row;
        private IList<ICell> cells;

        public ExcelRow()
        {
            row = new Row();
            cells = new List<ICell>();
        }

        internal ExcelRow(Row row)
        {
            this.row = row;
            if (row.ChildElements.Count == 0)
            {
                cells = new List<ICell>();
            } else
            {
                var cells = from cell in row.ChildElements as IEnumerable<Cell>
                            select new ExcelCell(cell);
                this.cells = (IList<ICell>)cells.ToList();
            }
        }

        public double? Height
        {
            get
            {
                var doubleValue = row.Height;
                if (doubleValue is null)
                {
                    return null;
                }

                return row.Height!.Value;
            }
            set
            {
                row.Height = value;
            }
        }
        public IList<ICell> Cells
        {
            get => cells;
            set
            {
                row.RemoveAllChildren<Row>();
                var excelCells = value as IList<ExcelCell>;
                if (excelCells != null)
                {
                    foreach (var cell in excelCells)
                    {
                        row.AppendChild(cell.cell);
                    }
                }
            }
        }
    }
}
